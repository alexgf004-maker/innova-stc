/**
 * kardex.js — Fase 1 (v4)
 * Formato de despacho alineado al documento físico real de DELSUR/INNOVA.
 *
 * Cambios v4:
 * - Formulario de salida con encabezado completo (usuario responsable, contratista,
 *   instalador, placa, fechas)
 * - Materiales con flag requiereSerial en catálogo
 * - Sección de seriales por material (individual o rango inicio/fin)
 * - Memo imprimible con formato del documento físico
 */

import {
  collection, doc, getDocs, addDoc, updateDoc,
  query, orderBy, where, serverTimestamp, increment
} from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';

import { showToast } from '../ui.js';

// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────
function safeNum(val) { const n = Number(val); return isNaN(n) ? 0 : n; }
function safeStr(val, fb = '—') {
  return (val !== undefined && val !== null && String(val).trim() !== '') ? String(val).trim() : fb;
}
function today() {
  return new Date().toISOString().split('T')[0];
}

// ─────────────────────────────────────────
// CONSTANTES DEL FORMULARIO DE DESPACHO
// Edita estas listas para actualizar las opciones del formulario.
// ─────────────────────────────────────────
const PLACAS = [
  'CPT-154',
  'CPT-156',
  'AU-250',
  'AU-200',
  'CNR-163',
  'P568DA',
  'P38DA6',
];

const USUARIOS_RESPONSABLES = [
  'NALVAR',
  'RGONZA',
  'JPEREZ',
];

const EMPRESAS_CONTRATISTAS = [
  'INNOVA',
];

// ─────────────────────────────────────────
// BADGE USUARIO OPERATIVO
// Aparece en todas las vistas de campo
// ─────────────────────────────────────────
function badgeUsuarioOperativo(session) {
  const u = session.usuarioOperativoAsignado;
  if (!u) return '';
  return '<div class="flex items-center gap-2 px-3 py-2 rounded-xl text-sm font-semibold text-white" style="background:#1B4F8A">' +
    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/></svg>' +
    'Trabajando con: ' + u +
  '</div>';
}

// ─────────────────────────────────────────
// INIT
// ─────────────────────────────────────────
export async function initKardex(session) {
  const container = document.getElementById('kardex-root');
  if (!container) return;
  const db = window.__firebase.db;

  // Campo sin asignación operativa — bloquear acceso
  if (session.role === 'campo' && !session.usuarioOperativoAsignado) {
    container.innerHTML =
      '<div class="flex flex-col items-center justify-center min-h-64 px-6 text-center space-y-4">' +
        '<div class="w-16 h-16 rounded-2xl flex items-center justify-center" style="background:#FEF2F2">' +
          '<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#C62828" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>' +
        '</div>' +
        '<div>' +
          '<p class="font-bold text-gray-900 text-lg">Sin asignación activa</p>' +
          '<p class="text-sm text-gray-500 mt-1">No tienes un usuario operativo asignado.</p>' +
          '<p class="text-sm text-gray-500">Contacta a administración para continuar.</p>' +
        '</div>' +
      '</div>';
    return;
  }

  renderShell(container, session);
  await showDashboard(db, session);
  bindNav(db, session);
}

// ─────────────────────────────────────────
// SHELL
// ─────────────────────────────────────────
function renderShell(container, session) {
  const canEdit  = ['admin','coordinadora'].includes(session.role);
  const isCampo  = session.role === 'campo';

  // Tabs por rol
  const tabsAdmin = `
    <button data-tab="dashboard"  class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Inicio</button>
    <button data-tab="inventario" class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Inventario</button>
    <button data-tab="historial"  class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Historial</button>
    <button data-tab="usuarios"   class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Usuarios</button>
    <button data-tab="solicitudes" class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Pedidos</button>`;

  const tabsCampo = `
    <button data-tab="dashboard"    class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Inicio</button>
    <button data-tab="consumo"      class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Consumo</button>
    <button data-tab="solicitar"    class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Solicitar</button>
    <button data-tab="mis-pedidos"  class="ktab flex-1 py-2 text-xs font-medium rounded-lg transition-colors">Pedidos</button>`;

  container.innerHTML =
    '<div class="space-y-4">' +
      '<div class="flex items-center justify-between">' +
        '<div>' +
          '<h1 class="text-xl font-semibold text-gray-900">Kardex</h1>' +
          '<p class="text-sm text-gray-500 mt-0.5">Control de materiales</p>' +
        '</div>' +
        (canEdit ? '<button id="btn-nueva-salida" class="inline-flex items-center gap-2 text-white text-sm font-medium px-4 py-2.5 rounded-lg" style="background-color:#1B4F8A"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg> Nueva salida</button>' : '') +
      '</div>' +
      '<div class="flex gap-1 bg-gray-100 rounded-xl p-1">' +
        (isCampo ? tabsCampo : tabsAdmin) +
      '</div>' +
      '<div id="kardex-content"></div>' +
    '</div>';
}

function setTab(tab) {
  document.querySelectorAll('.ktab').forEach(b => {
    const a = b.dataset.tab === tab;
    b.style.backgroundColor = a ? 'white' : 'transparent';
    b.style.color            = a ? '#1B4F8A' : '#6B7280';
    b.style.boxShadow        = a ? '0 1px 3px rgba(0,0,0,0.1)' : 'none';
  });
}

function bindNav(db, session) {
  document.querySelectorAll('.ktab').forEach(b => {
    b.addEventListener('click', async () => {
      const t = b.dataset.tab;
      if (t === 'dashboard')   await showDashboard(db, session);
      if (t === 'inventario')  await showInventario(db, session);
      if (t === 'historial')   await showHistorial(db, session);
      if (t === 'usuarios')    await showStockUsuarios(db, session);
      if (t === 'solicitudes') await showSolicitudes(db, session);
      if (t === 'consumo')     await showMisConsumos(db, session);
      if (t === 'solicitar')   await showSolicitarMaterial(db, session);
      if (t === 'mis-pedidos') await showMisSolicitudes(db, session);
    });
  });
  document.getElementById('btn-nueva-salida')?.addEventListener('click', () => showFormSalida(db, session));
}

function esValido(item) {
  const n = safeStr(item.name,'') || safeStr(item.nombre,'');
  const u = safeStr(item.unit,'') || safeStr(item.unidad,'');
  return n !== '' && u !== '';
}

function normalizeItem(raw) {
  return {
    ...raw,
    name:            safeStr(raw.name,'')    || safeStr(raw.nombre,''),
    unit:            safeStr(raw.unit,'')    || safeStr(raw.unidad,''),
    sapCode:         safeStr(raw.sapCode,''),
    axCode:          safeStr(raw.axCode,''),
    stock:           safeNum(raw.stock),
    minStock:        safeNum(raw.minStock || raw.stockMinimo || 5),
    requiereSerial:  raw.requiereSerial === true,
  };
}

// ─────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────
async function showDashboard(db, session) {
  setTab('dashboard');
  if (session.role === 'campo') {
    await showDashboardCampo(db, session);
    return;
  }
  await showDashboardAdmin(db, session);
}

async function showDashboardAdmin(db, session) {
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  try {
    const snap  = await getDocs(collection(db, 'kardex/inventario/items'));
    const items = snap.docs.map(d => normalizeItem({ id: d.id, ...d.data() })).filter(esValido);
    const total    = items.length;
    const sinStock = items.filter(i => safeNum(i.stock) <= 0).length;
    const alertas  = items.filter(i => safeNum(i.stock) <= safeNum(i.minStock || 5));

    const qS      = query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'));
    const snapS   = await getDocs(qS);
    const ultimas = snapS.docs.slice(0, 5).map(d => ({ id: d.id, ...d.data() }));

    content.innerHTML = `
      <div class="space-y-4">
        <div class="grid grid-cols-2 gap-3">
          ${statCard(total,    'Materiales', '#2196F3')}
          ${statCard(sinStock, 'Sin stock',  '#C62828')}
        </div>
        <div class="bg-white rounded-xl border border-gray-200 overflow-hidden">
          <div class="px-4 py-3 border-b border-gray-100 flex items-center justify-between">
            <p class="font-semibold text-sm text-gray-900">Últimas salidas</p>
            <button data-tab="historial" class="ktab text-xs font-medium" style="color:#2196F3">Ver todo</button>
          </div>
          ${ultimas.length === 0
            ? '<div class="px-4 py-6 text-center text-sm text-gray-400">Sin movimientos aún</div>'
            : '<div class="divide-y divide-gray-100">' + ultimas.map(s => `
              <div class="px-4 py-3">
                <div class="flex items-start justify-between gap-2">
                  <div class="min-w-0">
                    <p class="text-sm font-medium text-gray-900 truncate">${safeStr(s.usuarioResponsable || s.tecnicoNombre)}</p>
                    <p class="text-xs text-gray-400 mt-0.5">${fmtDate(s.fecha)}</p>
                  </div>
                  <p class="text-xs text-gray-500 shrink-0">${(s.items||[]).length} ítem(s)</p>
                </div>
                ${s.motivo ? `<p class="text-xs text-gray-400 mt-1 truncate">${s.motivo}</p>` : ''}
              </div>`).join('') + '</div>'
          }
        </div>
        ${alertas.length > 0 ? `
        <div class="bg-white rounded-xl border border-orange-200 overflow-hidden">
          <div class="px-4 py-3 border-b border-orange-100">
            <p class="font-semibold text-sm text-orange-800">⚠️ Atención al inventario</p>
          </div>
          <div class="divide-y divide-gray-100">
            ${alertas.map(i => {
              const stock = safeNum(i.stock);
              const bs = stock <= 0 ? 'background:#FEE2E2;color:#C62828' : 'background:#FEF3C7;color:#E65100';
              return `<div class="px-4 py-3 flex items-center justify-between gap-2">
                <div class="min-w-0">
                  <p class="text-sm text-gray-900 truncate">${safeStr(i.name)}</p>
                  <p class="text-xs text-gray-400">${safeStr(i.unit,'')}</p>
                </div>
                <span class="text-xs font-semibold px-2 py-1 rounded-full shrink-0" style="${bs}">
                  ${stock <= 0 ? 'Sin stock' : stock}
                </span>
              </div>`;
            }).join('')}
          </div>
        </div>` : ''}
      </div>`;

    document.querySelectorAll('.ktab').forEach(b => {
      b.addEventListener('click', async () => {
        const t = b.dataset.tab;
        if (t === 'dashboard')   await showDashboard(db, session);
        if (t === 'inventario')  await showInventario(db, session);
        if (t === 'historial')   await showHistorial(db, session);
        if (t === 'usuarios')    await showStockUsuarios(db, session);
        if (t === 'solicitudes') await showSolicitudes(db, session);
        if (t === 'consumo')     await showMisConsumos(db, session);
        if (t === 'solicitar')   await showSolicitarMaterial(db, session);
        if (t === 'mis-pedidos') await showMisSolicitudes(db, session);
      });
    });
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function statCard(val, label, color) {
  return '<div class="bg-white rounded-xl border border-gray-200 p-4 text-center">' +
    '<p class="text-3xl font-bold" style="color:' + color + '">' + val + '</p>' +
    '<p class="text-xs text-gray-500 mt-1">' + label + '</p>' +
  '</div>';
}

// ─────────────────────────────────────────
// DASHBOARD CAMPO
// Muestra: usuario operativo, su stock, solicitudes recientes
// ─────────────────────────────────────────
async function showDashboardCampo(db, session) {
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  const usuario = session.usuarioOperativoAsignado;

  try {
    // Stock del usuario operativo desde movimientos
    const snapSalidas = await getDocs(collection(db, 'kardex/movimientos/salidas'));
    const snapItems = await getDocs(collection(db, 'kardex/inventario/items'));

    const itemMap = {};
    snapItems.docs.forEach(function(d) {
      const item = normalizeItem({ id: d.id, ...d.data() });
      if (esValido(item)) itemMap[d.id] = item;
    });
    window.__kardexItemMap = itemMap;

    // Stock = salidas recibidas - consumos registrados
    const stockU = {};
    snapSalidas.docs.forEach(function(d) {
      const s = d.data();
      if ((s.usuarioResponsable || s.tecnicoNombre) !== usuario) return;
      (s.items || []).forEach(function(i) {
        const cant = safeNum(i.cantidad);
        if (!i.itemId || cant <= 0) return;
        stockU[i.itemId] = (stockU[i.itemId] || 0) + cant;
      });
    });

    // Descontar consumos del usuario
    try {
      const snapConsumos = await getDocs(query(
        collection(db, 'kardex/consumos'),
        where('usuarioOperativo', '==', usuario)
      ));
      snapConsumos.docs.forEach(function(d) {
        const c = d.data();
        (c.items || []).forEach(function(i) {
          const cant = safeNum(i.cantidad);
          if (!i.itemId || cant <= 0) return;
          stockU[i.itemId] = Math.max(0, (stockU[i.itemId] || 0) - cant);
        });
      });
    } catch(ce) { console.warn('No se pudieron cargar consumos:', ce); }

    if (!window.__kardexStockUsuario) window.__kardexStockUsuario = {};
    window.__kardexStockUsuario[usuario] = stockU;

    const misItems = Object.entries(stockU)
      .map(function(e) { return { id: e[0], cant: e[1], item: itemMap[e[0]] }; })
      .filter(function(e) { return e.cant > 0 && e.item; })
      .sort(function(a, b) { return safeStr(a.item.name).localeCompare(safeStr(b.item.name)); });

    // Solicitudes recientes del usuario (sin orderBy para evitar índice compuesto)
    let solRecientes = [];
    let solPendientes = 0;
    try {
      const snapSol = await getDocs(query(
        collection(db, 'solicitudes_material'),
        where('usuarioUid', '==', session.uid)
      ));
      const todas = snapSol.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });
      todas.sort(function(a, b) {
        const fa = a.fecha?.seconds || 0;
        const fb = b.fecha?.seconds || 0;
        return fb - fa;
      });
      solRecientes  = todas.slice(0, 3);
      solPendientes = todas.filter(function(d) { return d.estado === 'pendiente'; }).length;
    } catch(solErr) { console.warn('No se pudieron cargar solicitudes:', solErr); }

    const ESTADO_BADGE = {
      pendiente: { label: 'Pendiente', bg: '#FEF3C7', color: '#B45309' },
      aprobado:  { label: 'Aprobado',  bg: '#DCFCE7', color: '#166534' },
      rechazado: { label: 'Rechazado', bg: '#FEE2E2', color: '#C62828' },
    };

    function tc(str) {
      return safeStr(str).toLowerCase().replace(/\w/g, function(c) { return c.toUpperCase(); });
    }

    // Render
    let html = '<div class="space-y-4">';

    // Badge usuario operativo — prominente
    html += '<div class="rounded-2xl p-4 text-white" style="background:#1B4F8A">' +
      '<p class="text-xs font-medium opacity-80 uppercase tracking-wider">Trabajando con</p>' +
      '<p class="text-2xl font-black mt-1">' + usuario + '</p>' +
      '<p class="text-xs opacity-70 mt-0.5">' + misItems.length + ' material' + (misItems.length !== 1 ? 'es' : '') + ' asignados</p>' +
    '</div>';

    // Acciones rápidas
    html += '<div class="grid grid-cols-2 gap-3">' +
      '<button id="db-consumo" class="flex items-center justify-between px-4 py-3.5 bg-white rounded-xl border-2 text-left active:opacity-80" style="border-color:#166534">' +
        '<div>' +
          '<p class="font-semibold text-gray-900 text-sm">Registrar consumo</p>' +
          '<p class="text-xs text-gray-400 mt-0.5">Material usado en OT</p>' +
        '</div>' +
        '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#166534" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>' +
      '</button>' +
      '<button id="db-solicitar" class="flex items-center justify-between px-4 py-3.5 bg-white rounded-xl border-2 border-blue-200 text-left active:bg-blue-50">' +
        '<div>' +
          '<p class="font-semibold text-gray-900 text-sm">Solicitar material</p>' +
          '<p class="text-xs text-gray-400 mt-0.5">Pide a bodega</p>' +
        '</div>' +
        '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#1B4F8A" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>' +
      '</button>' +
    '</div>';

    // Mi material
    if (misItems.length > 0) {
      html += '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
        '<div class="px-4 py-3 border-b border-gray-100 flex items-center justify-between">' +
          '<p class="font-semibold text-sm text-gray-900">Mi material</p>' +
          '<span class="text-xs text-gray-400">' + misItems.length + ' ítem' + (misItems.length !== 1 ? 's' : '') + '</span>' +
        '</div>' +
        '<div class="px-3 py-2 space-y-1.5">' +
        misItems.slice(0, 8).map(function(e) {
          const sap = safeStr(e.item.sapCode, '');
          const unit = safeStr(e.item.unit, '');
          return '<div class="flex items-center gap-3 px-3 py-3 rounded-xl bg-gray-50">' +
            // Cantidad — destacada a la izquierda
            '<div class="shrink-0 w-14 text-center">' +
              '<p class="text-2xl font-black leading-none" style="color:#1B4F8A">' + e.cant + '</p>' +
              '<p class="text-xs text-gray-400 mt-0.5">' + unit + '</p>' +
            '</div>' +
            // Separador vertical
            '<div class="w-px self-stretch bg-gray-200 shrink-0"></div>' +
            // Nombre y SAP
            '<div class="flex-1 min-w-0">' +
              '<p class="text-sm font-semibold text-gray-900 leading-snug" style="display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden">' + tc(e.item.name) + '</p>' +
              (sap ? '<p class="text-xs text-gray-400 font-mono mt-0.5">SAP ' + sap + '</p>' : '') +
            '</div>' +
          '</div>';
        }).join('') +
        (misItems.length > 8 ? '<p class="text-xs text-gray-400 text-center py-2">+' + (misItems.length - 8) + ' materiales más</p>' : '') +
        '</div>' +
      '</div>';
    }

    // Solicitudes recientes
    if (solRecientes.length > 0) {
      html += '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
        '<div class="px-4 py-3 border-b border-gray-100 flex items-center justify-between">' +
          '<p class="font-semibold text-sm text-gray-900">Mis solicitudes</p>' +
          (solPendientes > 0 ? '<span class="text-xs font-bold px-2 py-0.5 rounded-full text-white" style="background:#B45309">' + solPendientes + ' pendiente' + (solPendientes !== 1 ? 's' : '') + '</span>' : '') +
        '</div>' +
        '<div class="divide-y divide-gray-50">' +
        solRecientes.map(function(s) {
          const badge = ESTADO_BADGE[s.estado] || ESTADO_BADGE.pendiente;
          return '<div class="px-4 py-3 flex items-center justify-between gap-2">' +
            '<div class="min-w-0">' +
              '<p class="text-xs text-gray-400">' + fmtDate(s.fecha) + '</p>' +
              '<p class="text-xs text-gray-600 truncate mt-0.5">' + (s.materiales || []).length + ' material(es)</p>' +
            '</div>' +
            '<span class="text-xs font-semibold px-2 py-0.5 rounded-full shrink-0" style="background:' + badge.bg + ';color:' + badge.color + '">' + badge.label + '</span>' +
          '</div>';
        }).join('') +
        '</div>' +
      '</div>';
    }

    html += '</div>';
    content.innerHTML = html;

    content.querySelector('#db-consumo')?.addEventListener('click', function() {
      showRegistrarConsumo(db, session);
    });
    content.querySelector('#db-solicitar')?.addEventListener('click', function() {
      showSolicitarMaterial(db, session);
      setTab('solicitar');
    });

    // Rebind tabs
    document.querySelectorAll('.ktab').forEach(function(b) {
      b.addEventListener('click', async function() {
        const t = b.dataset.tab;
        if (t === 'dashboard')   await showDashboard(db, session);
        if (t === 'solicitar')   await showSolicitarMaterial(db, session);
        if (t === 'mis-pedidos') await showMisSolicitudes(db, session);
      });
    });

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// INVENTARIO
// ─────────────────────────────────────────
async function showInventario(db, session) {
  setTab('inventario');
  const content  = document.getElementById('kardex-content');
  content.innerHTML = loading();
  const canEdit = ['admin','coordinadora'].includes(session.role);

  let seleccionMode = false;
  let seleccionados = new Set();

  try {
    const snap     = await getDocs(collection(db, 'kardex/inventario/items'));
    const allItems = snap.docs.map(d => normalizeItem({ id: d.id, ...d.data() })).filter(esValido)
      .sort((a,b) => safeStr(a.name).localeCompare(safeStr(b.name)));

    content.innerHTML = `
      <div class="space-y-3">
        <div id="inv-bar-normal" class="flex gap-2 justify-end">
          ${session.role === 'admin' ? `
          <button id="btn-importar-excel"
            class="inline-flex items-center gap-2 text-sm font-medium px-3 py-2 rounded-lg border border-gray-300 text-gray-700 hover:bg-gray-50">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
              <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
            </svg>
            Importar
          </button>
          <button id="btn-seleccionar"
            class="inline-flex items-center gap-2 text-sm font-medium px-3 py-2 rounded-lg border border-gray-300 text-gray-700 hover:bg-gray-50">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/>
            </svg>
            Seleccionar
          </button>` : ''}
          ${canEdit ? `
          <button id="btn-nuevo-item"
            class="inline-flex items-center gap-2 text-white text-sm font-medium px-3 py-2 rounded-lg"
            style="background-color:#1B4F8A">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
            </svg>
            Agregar
          </button>` : ''}
        </div>

        <div id="inv-bar-seleccion" class="hidden flex items-center justify-between bg-blue-50 border border-blue-200 rounded-xl px-4 py-2.5">
          <span id="inv-sel-count" class="text-sm font-medium text-blue-800">0 seleccionados</span>
          <div class="flex gap-2">
            <button id="btn-cancelar-sel" class="text-sm font-medium px-3 py-1.5 rounded-lg border border-gray-300 text-gray-700 bg-white">Cancelar</button>
            <button id="btn-borrar-sel" class="text-sm font-medium px-3 py-1.5 rounded-lg text-white disabled:opacity-40" style="background:#C62828" disabled>Eliminar</button>
          </div>
        </div>

        <div class="relative">
          <svg class="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input id="inv-search" type="text" placeholder="Buscar por nombre, SAP o AX..."
            class="w-full border border-gray-300 rounded-xl pl-9 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>

        <div id="inv-tabla" class="bg-white rounded-xl border border-gray-200 overflow-hidden"></div>
      </div>`;

    function updateSelCount() {
      const countEl = document.getElementById('inv-sel-count');
      const btnDel  = document.getElementById('btn-borrar-sel');
      if (countEl) countEl.textContent = `${seleccionados.size} seleccionado(s)`;
      if (btnDel)  btnDel.disabled = seleccionados.size === 0;
    }

    function renderTabla(items) {
      const tabla = document.getElementById('inv-tabla');
      if (!tabla) return;
      if (items.length === 0) {
        tabla.innerHTML = `<div class="py-12 text-center text-sm text-gray-400">Sin resultados.</div>`;
        return;
      }
      tabla.innerHTML = `
        <div class="overflow-x-auto">
          <table class="w-full text-sm">
            <thead class="bg-gray-50 border-b border-gray-200">
              <tr>
                ${seleccionMode ? '<th class="px-4 py-3 w-8"></th>' : ''}
                <th class="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Material</th>
                <th class="text-center px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide w-20">Stock</th>
                ${canEdit && !seleccionMode ? '<th class="px-4 py-3 w-20"></th>' : ''}
              </tr>
            </thead>
            <tbody class="divide-y divide-gray-100">
              ${items.map(item => itemRow(item, canEdit, seleccionMode, seleccionados.has(item.id))).join('')}
            </tbody>
          </table>
        </div>`;

      if (seleccionMode) {
        tabla.querySelectorAll('[data-sel]').forEach(cb => {
          cb.addEventListener('change', () => {
            if (cb.checked) seleccionados.add(cb.dataset.sel);
            else seleccionados.delete(cb.dataset.sel);
            updateSelCount();
            cb.closest('tr').style.background = cb.checked ? '#EFF6FF' : '';
          });
          if (seleccionados.has(cb.dataset.sel)) {
            cb.checked = true;
            cb.closest('tr').style.background = '#EFF6FF';
          }
        });
      } else if (canEdit) {
        tabla.querySelectorAll('[data-entrada]').forEach(b => {
          b.addEventListener('click', () => showFormEntrada(db, session, allItems.find(i => i.id === b.dataset.entrada)));
        });
        tabla.querySelectorAll('[data-edit]').forEach(b => {
          b.addEventListener('click', () => showFormItem(db, session, allItems.find(i => i.id === b.dataset.edit)));
        });
        tabla.querySelectorAll('[data-seriales]').forEach(b => {
          b.addEventListener('click', () => showSeriales(db, session, allItems.find(i => i.id === b.dataset.seriales)));
        });
      }
    }

    renderTabla(allItems);

    document.getElementById('inv-search')?.addEventListener('input', (e) => {
      const q = e.target.value.trim().toLowerCase();
      if (!q) { renderTabla(allItems); return; }
      renderTabla(allItems.filter(i =>
        safeStr(i.name).toLowerCase().includes(q) ||
        safeStr(i.sapCode,'').toLowerCase().includes(q) ||
        safeStr(i.axCode,'').toLowerCase().includes(q)
      ));
    });

    if (canEdit) {
      document.getElementById('btn-nuevo-item')?.addEventListener('click', () => showFormItem(db, session, null));
    }
    if (session.role === 'admin') {
      document.getElementById('btn-importar-excel')?.addEventListener('click', () => showImportExcel(db, session, allItems, renderTabla));
      document.getElementById('btn-seleccionar')?.addEventListener('click', () => {
        seleccionMode = true; seleccionados.clear();
        document.getElementById('inv-bar-normal').classList.add('hidden');
        document.getElementById('inv-bar-seleccion').classList.remove('hidden');
        renderTabla(allItems); updateSelCount();
      });
      document.getElementById('btn-cancelar-sel')?.addEventListener('click', () => {
        seleccionMode = false; seleccionados.clear();
        document.getElementById('inv-bar-seleccion').classList.add('hidden');
        document.getElementById('inv-bar-normal').classList.remove('hidden');
        renderTabla(allItems);
      });
      document.getElementById('btn-borrar-sel')?.addEventListener('click', async () => {
        if (seleccionados.size === 0) return;
        await showBorrarSeleccionados(db, session, allItems.filter(i => seleccionados.has(i.id)), async () => {
          seleccionMode = false; seleccionados.clear();
          await showInventario(db, session);
        });
      });
    }
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function itemRow(item, canEdit, selMode = false, isSelected = false) {
  const stock  = safeNum(item.stock);
  const min    = safeNum(item.minStock || 5);
  const unit   = safeStr(item.unit, '');
  const name   = safeStr(item.name, '—');
  const sap    = safeStr(item.sapCode, '');
  const ax     = safeStr(item.axCode, '');

  const badgeStyle = stock <= 0
    ? 'background:#FEE2E2;color:#C62828'
    : stock <= min ? 'background:#FEF3C7;color:#E65100'
    : 'background:#DCFCE7;color:#166534';

  const codigos = [
    sap ? `<span class="font-mono">SAP ${sap}</span>` : '',
    ax  ? `<span class="font-mono">AX ${ax}</span>`   : '',
  ].filter(Boolean).join('<span class="mx-1 text-gray-300">·</span>');

  const serialBadge = item.requiereSerial
    ? `<span class="text-xs px-1.5 py-0.5 rounded font-medium" style="background:#EFF6FF;color:#1D4ED8">Serial</span>`
    : '';

  return `
    <tr class="hover:bg-gray-50 transition-colors">
      ${selMode ? `<td class="px-3 py-3">
        <input type="checkbox" data-sel="${item.id}" class="w-4 h-4 rounded border-gray-300 cursor-pointer" style="accent-color:#1B4F8A" ${isSelected ? 'checked' : ''}/>
      </td>` : ''}
      <td class="px-4 py-3">
        <div class="flex items-start gap-2">
          <div class="min-w-0">
            <p class="font-medium text-gray-900 leading-tight">${name} ${serialBadge}</p>
            ${codigos ? `<p class="text-xs text-gray-400 mt-0.5">${codigos}</p>` : ''}
          </div>
        </div>
      </td>
      <td class="px-4 py-3 text-center">
        <span class="text-sm font-bold px-2.5 py-1 rounded-full" style="${badgeStyle}">${stock}</span>
        <p class="text-xs text-gray-400 mt-0.5">${unit}</p>
      </td>
      ${canEdit && !selMode ? `
      <td class="px-4 py-3">
        <div class="flex items-center justify-end gap-1">
          <button data-entrada="${item.id}" title="Registrar entrada" class="p-1.5 rounded text-gray-400 hover:text-green-600 transition-colors">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12l7 7 7-7"/></svg>
          </button>
          ${item.requiereSerial ? `<button data-seriales="${item.id}" title="Ver seriales" class="p-1.5 rounded text-gray-400 hover:text-purple-600 transition-colors">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="9" y1="9" x2="15" y2="9"/><line x1="9" y1="12" x2="15" y2="12"/><line x1="9" y1="15" x2="12" y2="15"/></svg>
          </button>` : ''}
          <button data-edit="${item.id}" title="Editar" class="p-1.5 rounded text-gray-400 hover:text-blue-500 transition-colors">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/>
              <path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/>
            </svg>
          </button>
        </div>
      </td>` : ''}
    </tr>`;
}

// ─────────────────────────────────────────
// VISTA DE SERIALES POR MATERIAL
// ─────────────────────────────────────────
async function showSeriales(db, session, item) {
  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<h2 class="font-semibold text-gray-900">Seriales</h2>' +
        '<p class="text-xs text-gray-400 mt-0.5">' + safeStr(item?.name) + '</p>' +
      '</div>' +
      '<button id="sv-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="px-5 pt-3 shrink-0">' +
      '<div class="relative mb-3">' +
        '<svg class="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>' +
        '<input id="sv-buscar" type="text" placeholder="Buscar serial..." class="w-full border border-gray-200 rounded-xl pl-9 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>' +
      '</div>' +
      '<div class="flex gap-2 mb-3">' +
        '<button id="sv-tab-disp" class="flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-blue-500 bg-blue-50 text-blue-700">Disponibles <span id="sv-cnt-disp" class="ml-1 text-xs bg-blue-200 text-blue-800 px-1.5 py-0.5 rounded-full">0</span></button>' +
        '<button id="sv-tab-desp" class="flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-gray-200 bg-white text-gray-500">Despachados <span id="sv-cnt-desp" class="ml-1 text-xs bg-gray-200 text-gray-600 px-1.5 py-0.5 rounded-full">0</span></button>' +
      '</div>' +
    '</div>' +
    '<div id="sv-lista" class="flex-1 overflow-y-auto px-5 pb-5 space-y-2">' +
      '<p class="text-sm text-gray-400 text-center py-8">Cargando...</p>' +
    '</div>'
  );

  ov.querySelector('#sv-close').onclick = () => ov.remove();

  let tabActual = 'disponible';
  let todosSeriales = [];
  let busq = '';

  function fmt(ts) {
    if (!ts) return '—';
    const d = ts.toDate ? ts.toDate() : new Date(ts);
    return d.toLocaleDateString('es-SV', { day:'2-digit', month:'short', year:'numeric' });
  }

  function renderLista() {
    const lista = ov.querySelector('#sv-lista');
    if (!lista) return;
    const q = busq.toLowerCase();
    const filtrados = todosSeriales.filter(s =>
      s.estado === tabActual &&
      (!q || s.serial.toLowerCase().includes(q))
    );

    if (!filtrados.length) {
      lista.innerHTML = '<p class="text-sm text-gray-400 text-center py-8">' +
        (q ? 'Sin resultados para "' + q + '"' : 'Sin seriales ' + (tabActual === 'disponible' ? 'disponibles' : 'despachados')) +
      '</p>';
      return;
    }

    lista.innerHTML = filtrados.map(function(s) {
      const isDisp = s.estado === 'disponible';
      const badgeBg    = isDisp ? '#DCFCE7' : '#FEE2E2';
      const badgeColor = isDisp ? '#166534' : '#C62828';
      const badgeText  = isDisp ? 'Disponible' : 'Despachado';

      return '<div class="bg-white border border-gray-200 rounded-2xl px-4 py-3">' +
        '<div class="flex items-center justify-between mb-1">' +
          '<span class="font-mono text-sm font-bold text-gray-900">' + s.serial + '</span>' +
          '<span class="text-xs font-semibold px-2 py-0.5 rounded-full" style="background:' + badgeBg + ';color:' + badgeColor + '">' + badgeText + '</span>' +
        '</div>' +
        '<div class="text-xs text-gray-400 space-y-0.5">' +
          '<p>📥 Entrada: ' + fmt(s.fechaEntrada) + '</p>' +
          (!isDisp ? '<p>📤 Salida: ' + fmt(s.fechaSalida) + '</p>' : '') +
          (!isDisp && s.usuarioDespacho ? '<p>👤 Despachado a: <span class="font-medium text-gray-600">' + s.usuarioDespacho + '</span></p>' : '') +
        '</div>' +
      '</div>';
    }).join('');
  }

  function updateTabs() {
    const disp = todosSeriales.filter(s => s.estado === 'disponible').length;
    const desp = todosSeriales.filter(s => s.estado === 'despachado').length;
    const cntD = ov.querySelector('#sv-cnt-disp');
    const cntDp = ov.querySelector('#sv-cnt-desp');
    if (cntD)  cntD.textContent  = disp;
    if (cntDp) cntDp.textContent = desp;

    const tabDisp = ov.querySelector('#sv-tab-disp');
    const tabDesp = ov.querySelector('#sv-tab-desp');
    if (tabActual === 'disponible') {
      tabDisp.className = 'flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-blue-500 bg-blue-50 text-blue-700';
      tabDesp.className = 'flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-gray-200 bg-white text-gray-500';
    } else {
      tabDesp.className = 'flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-blue-500 bg-blue-50 text-blue-700';
      tabDisp.className = 'flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-gray-200 bg-white text-gray-500';
    }
  }

  // Cargar seriales
  try {
    const snap = await getDocs(query(
      collection(db, 'kardex/seriales/items'),
      where('itemId', '==', item.id)
    ));
    todosSeriales = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    todosSeriales.sort((a, b) => a.serial.localeCompare(b.serial, undefined, { numeric: true }));
    updateTabs();
    renderLista();
  } catch(e) {
    ov.querySelector('#sv-lista').innerHTML = '<p class="text-sm text-red-500 text-center py-8">Error al cargar seriales.</p>';
    console.error(e);
  }

  ov.querySelector('#sv-buscar').addEventListener('input', function(e) {
    busq = e.target.value.trim();
    renderLista();
  });
  ov.querySelector('#sv-tab-disp').addEventListener('click', function() {
    tabActual = 'disponible'; updateTabs(); renderLista();
  });
  ov.querySelector('#sv-tab-desp').addEventListener('click', function() {
    tabActual = 'despachado'; updateTabs(); renderLista();
  });
}

// ─────────────────────────────────────────
// FORM — AGREGAR / EDITAR MATERIAL
// ─────────────────────────────────────────
function showFormItem(db, session, item) {
  const esNuevo = !item;
  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">${esNuevo ? 'Nuevo material' : 'Editar material'}</h2>
      <button id="fi-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div class="px-4 py-3 space-y-3">
      <div>
        <label class="block text-xs font-medium text-gray-600 mb-1">Nombre del material</label>
        <input id="fi-name" type="text" value="${safeStr(item?.name||item?.nombre,'')}"
          placeholder="Ej. MEDIDOR MONOFÁSICO 120V"
          class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1">Código SAP</label>
          <input id="fi-sap" type="text" value="${safeStr(item?.sapCode,'')}" placeholder="221477"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1">Código AX</label>
          <input id="fi-ax" type="text" value="${safeStr(item?.axCode,'')}" placeholder="50203"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1">Unidad</label>
          <input id="fi-unit" type="text" value="${safeStr(item?.unit||item?.unidad,'')}" placeholder="unidad, m, kg"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1">Stock mínimo</label>
          <input id="fi-min" type="number" min="0" value="${safeNum(item?.minStock||item?.stockMinimo) || 5}"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
      </div>
      <label class="flex items-center gap-2 cursor-pointer">
        <input id="fi-serial" type="checkbox" ${item?.requiereSerial ? 'checked' : ''} class="w-4 h-4 rounded" style="accent-color:#1B4F8A"/>
        <span class="text-sm text-gray-700">Requiere registro de seriales / sellos</span>
      </label>
      ${esNuevo ? `
      <div class="bg-blue-50 border border-blue-100 rounded-lg p-2.5">
        <p class="text-xs font-medium text-blue-800 mb-1">Stock inicial <span class="font-normal text-blue-500">(entrada inicial)</span></p>
        <input id="fi-stock" type="number" min="0" value="0"
          class="w-full border border-blue-200 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-400 text-center font-semibold"/>
      </div>` : `
      <p class="text-xs text-gray-400">Para cambiar el stock usa el botón de entrada ↓ en la lista.</p>`}
      <div id="fi-err" class="hidden text-xs text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2"></div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fi-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cancelar</button>
      <button id="fi-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">
        ${esNuevo ? 'Agregar' : 'Guardar cambios'}
      </button>
    </div>`);

  ov.querySelector('#fi-close').onclick = ov.querySelector('#fi-cancel').onclick = () => ov.remove();
  ov.querySelector('#fi-submit').onclick = async () => {
    const name          = ov.querySelector('#fi-name').value.trim();
    const sapCode       = ov.querySelector('#fi-sap').value.trim();
    const axCode        = ov.querySelector('#fi-ax').value.trim();
    const unit          = ov.querySelector('#fi-unit').value.trim();
    const minStock      = safeNum(ov.querySelector('#fi-min').value);
    const requiereSerial = ov.querySelector('#fi-serial').checked;
    const errEl         = ov.querySelector('#fi-err');
    const btn           = ov.querySelector('#fi-submit');

    errEl.classList.add('hidden');
    if (!name) { errEl.textContent = 'El nombre es obligatorio.'; errEl.classList.remove('hidden'); return; }
    if (!unit) { errEl.textContent = 'La unidad es obligatoria.'; errEl.classList.remove('hidden'); return; }

    btn.disabled = true; btn.textContent = 'Verificando...';
    try {
      if (sapCode || axCode) {
        const snapAll  = await getDocs(collection(db, 'kardex/inventario/items'));
        const existing = snapAll.docs.map(d => ({ id: d.id, ...d.data() })).filter(esValido);
        for (const ex of existing) {
          if (item && ex.id === item.id) continue;
          if (sapCode && safeStr(ex.sapCode,'') === sapCode) {
            errEl.textContent = `SAP ${sapCode} ya existe: ${safeStr(ex.name||ex.nombre)}`;
            errEl.classList.remove('hidden');
            btn.disabled = false; btn.textContent = esNuevo ? 'Agregar' : 'Guardar cambios';
            return;
          }
          if (axCode && safeStr(ex.axCode,'') === axCode) {
            errEl.textContent = `AX ${axCode} ya existe: ${safeStr(ex.name||ex.nombre)}`;
            errEl.classList.remove('hidden');
            btn.disabled = false; btn.textContent = esNuevo ? 'Agregar' : 'Guardar cambios';
            return;
          }
        }
      }
      btn.textContent = 'Guardando...';
      if (esNuevo) {
        const stockInicial = safeNum(ov.querySelector('#fi-stock').value);
        const ref = await addDoc(collection(db, 'kardex/inventario/items'), {
          name, sapCode, axCode, unit, stock: stockInicial, minStock, requiereSerial,
          creadoEn: serverTimestamp(), creadoPor: session.uid,
        });
        if (stockInicial > 0) {
          await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
            itemId: ref.id, itemNombre: name, tipo: 'entrada_inicial',
            cantidad: stockInicial, motivo: 'Stock inicial',
            stockAntes: 0, stockDespues: stockInicial,
            fecha: serverTimestamp(), registradoPor: session.uid, registradoPorNombre: session.displayName,
          });
        }
      } else {
        await updateDoc(doc(db, 'kardex/inventario/items', item.id), {
          name, sapCode, axCode, unit, minStock, requiereSerial,
        });
      }
      ov.remove();
      showToast(`Material ${esNuevo ? 'agregado' : 'actualizado'}.`, 'success');
      await showInventario(db, session);
    } catch(e) {
      errEl.textContent = 'Error al guardar.';
      errEl.classList.remove('hidden');
      btn.disabled = false; btn.textContent = esNuevo ? 'Agregar' : 'Guardar cambios';
      console.error(e);
    }
  };
}

// ─────────────────────────────────────────
// FORM — ENTRADA DE MATERIAL
// ─────────────────────────────────────────
function showFormEntrada(db, session, item) {
  const esSello  = item?.sapCode === '354549';
  const esSer    = !!item?.requiereSerial;
  let   modoSer  = esSello ? 'rango' : 'individual';

  const serBlock = !esSer ? '' :
    '<div class="border-t border-gray-100 pt-4 space-y-3">' +
      '<p class="text-sm font-semibold text-gray-700">Seriales que ingresan</p>' +
      (!esSello ?
        '<div class="flex gap-2">' +
          '<button id="fe-modo-ind" class="flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-blue-500 bg-blue-50 text-blue-700">Individual</button>' +
          '<button id="fe-modo-rng" class="flex-1 py-2 rounded-xl text-sm font-semibold border-2 border-gray-200 bg-white text-gray-500">Rango</button>' +
        '</div>' : '') +
      '<div id="fe-ser-campos">' +
        '<p class="text-xs text-gray-400 mb-1">Seriales (uno por línea)</p>' +
        '<textarea id="fe-sers" rows="5" placeholder="12345001&#10;12345002&#10;12345003" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400 resize-none"></textarea>' +
      '</div>' +
    '</div>';

  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">' +
      '<h2 class="font-semibold text-gray-900">Registrar entrada</h2>' +
      '<button id="fe-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="px-5 py-4 space-y-4" style="max-height:80dvh;overflow-y:auto">' +
      '<div class="bg-gray-50 rounded-xl p-3">' +
        '<p class="text-sm font-medium text-gray-900">' + safeStr(item?.name) + '</p>' +
        '<p class="text-xs text-gray-400 font-mono mt-0.5">' +
          (item?.sapCode ? 'SAP ' + item.sapCode : '') +
          (item?.sapCode && item?.axCode ? ' · ' : '') +
          (item?.axCode ? 'AX ' + item.axCode : '') +
        '</p>' +
        '<p class="text-xs text-gray-500 mt-1">Stock actual: <strong>' + safeNum(item?.stock) + ' ' + safeStr(item?.unit,'') + '</strong></p>' +
      '</div>' +
      (!esSer ?
        '<div>' +
          '<label class="block text-sm font-medium text-gray-700 mb-1.5">Cantidad que ingresa</label>' +
          '<input id="fe-cant" type="number" min="1" placeholder="0" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 text-center text-lg font-semibold"/>' +
        '</div>'
        :
        '<div class="bg-blue-50 border border-blue-100 rounded-xl p-3 text-center">' +
          '<p class="text-xs text-blue-500 mb-0.5">Cantidad calculada del serial</p>' +
          '<p id="fe-cant-display" class="text-3xl font-black text-blue-700">—</p>' +
          '<input id="fe-cant" type="hidden" value="0"/>' +
        '</div>'
      ) +
      '<div>' +
        '<label class="block text-sm font-medium text-gray-700 mb-1.5">Motivo</label>' +
        '<input id="fe-motivo" type="text" placeholder="Ej. Compra, Reposición" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>' +
      '</div>' +
      serBlock +
      '<div id="fe-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex gap-3">' +
      '<button id="fe-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm">Cancelar</button>' +
      '<button id="fe-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#2E7D32">Registrar entrada</button>' +
    '</div>');

  ov.querySelector('#fe-close').onclick = ov.querySelector('#fe-cancel').onclick = () => ov.remove();

  // Modo seriales (solo medidores)
  function actualizarCantDisplay(cant) {
    const disp = ov.querySelector('#fe-cant-display');
    const inp  = ov.querySelector('#fe-cant');
    if (disp) disp.textContent = cant > 0 ? cant : '—';
    if (inp)  inp.value = cant;
  }

  function actualizarModo() {
    const campos = ov.querySelector('#fe-ser-campos');
    if (!campos) return;
    const indBtn = ov.querySelector('#fe-modo-ind');
    const rngBtn = ov.querySelector('#fe-modo-rng');
    if (indBtn) {
      indBtn.className = 'flex-1 py-2 rounded-xl text-sm font-semibold border-2 ' +
        (modoSer==='individual' ? 'border-blue-500 bg-blue-50 text-blue-700' : 'border-gray-200 bg-white text-gray-500');
      rngBtn.className = 'flex-1 py-2 rounded-xl text-sm font-semibold border-2 ' +
        (modoSer==='rango' ? 'border-blue-500 bg-blue-50 text-blue-700' : 'border-gray-200 bg-white text-gray-500');
    }
    if (modoSer === 'rango') {
      campos.innerHTML =
        '<div class="flex gap-2">' +
          '<div class="flex-1"><p class="text-xs text-gray-400 mb-1">Serial inicio</p>' +
            '<input id="fe-ini" type="text" placeholder="Ej: 12345001" class="w-full border border-gray-300 rounded-lg px-3 py-2.5 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>' +
          '<div class="flex-1"><p class="text-xs text-gray-400 mb-1">Serial fin</p>' +
            '<input id="fe-fin" type="text" placeholder="Ej: 12345030" class="w-full border border-gray-300 rounded-lg px-3 py-2.5 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>' +
        '</div>';
      // Auto-calcular cantidad del rango
      setTimeout(() => {
        const calcRango = () => {
          const ini = (ov.querySelector('#fe-ini')?.value || '').trim();
          const fin = (ov.querySelector('#fe-fin')?.value || '').trim();
          const nIni = parseInt(ini.replace(/\D/g,''), 10);
          const nFin = parseInt(fin.replace(/\D/g,''), 10);
          const cantCalc = (!isNaN(nIni) && !isNaN(nFin) && nFin >= nIni) ? (nFin - nIni + 1) : 0;
          actualizarCantDisplay(cantCalc);
        };
        ov.querySelector('#fe-ini')?.addEventListener('input', calcRango);
        ov.querySelector('#fe-fin')?.addEventListener('input', calcRango);
      }, 50);
    } else {
      campos.innerHTML =
        '<p class="text-xs text-gray-400 mb-1">Seriales (uno por línea)</p>' +
        '<textarea id="fe-sers" rows="5" placeholder="12345001&#10;12345002&#10;12345003" class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400 resize-none"></textarea>';
      // Auto-calcular cantidad del textarea
      setTimeout(() => {
        ov.querySelector('#fe-sers')?.addEventListener('input', function() {
          const sers = this.value.trim().split('\n').map(s=>s.trim()).filter(Boolean);
          actualizarCantDisplay(sers.length);
        });
      }, 50);
    }
    actualizarCantDisplay(0);
  }

  ov.querySelector('#fe-modo-ind')?.addEventListener('click', () => { modoSer='individual'; actualizarModo(); });
  ov.querySelector('#fe-modo-rng')?.addEventListener('click', () => { modoSer='rango'; actualizarModo(); });
  // Inicializar modo actual
  if (esSer) actualizarModo();

  ov.querySelector('#fe-submit').onclick = async () => {
    const motivo = ov.querySelector('#fe-motivo').value.trim();
    const errEl  = ov.querySelector('#fe-err');
    const btn    = ov.querySelector('#fe-submit');
    errEl.classList.add('hidden');
    // Recolectar seriales primero para calcular cantidad real
    let seriales = [], serialInicio = '', serialFin = '';
    let cant = 0;
    if (esSer) {
      if (modoSer === 'rango') {
        serialInicio = (ov.querySelector('#fe-ini')?.value || '').trim();
        serialFin    = (ov.querySelector('#fe-fin')?.value || '').trim();
        if (!serialInicio || !serialFin) {
          errEl.textContent = 'Ingresa el serial de inicio y fin.';
          errEl.classList.remove('hidden'); return;
        }
        const nIni = parseInt(serialInicio.replace(/\D/g,''), 10);
        const nFin = parseInt(serialFin.replace(/\D/g,''), 10);
        cant = (!isNaN(nIni) && !isNaN(nFin) && nFin >= nIni) ? (nFin - nIni + 1) : 0;
        if (cant <= 0) { errEl.textContent = 'El rango es inválido.'; errEl.classList.remove('hidden'); return; }
      } else {
        const raw = (ov.querySelector('#fe-sers')?.value || '').trim();
        seriales = raw ? raw.split('\n').map(s => s.trim()).filter(Boolean) : [];
        if (!seriales.length) { errEl.textContent = 'Ingresa al menos un serial.'; errEl.classList.remove('hidden'); return; }
        cant = seriales.length;
      }
    } else {
      cant = safeNum(ov.querySelector('#fe-cant').value);
      if (cant <= 0) { errEl.textContent = 'Cantidad inválida.'; errEl.classList.remove('hidden'); return; }
    }

    btn.disabled = true; btn.textContent = 'Registrando...';
    try {
      const stockAntes = safeNum(item?.stock);
      await updateDoc(doc(db, 'kardex/inventario/items', item.id), { stock: increment(cant) });
      await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
        itemId: item.id, itemNombre: safeStr(item?.name), tipo: 'entrada', cantidad: cant, motivo,
        stockAntes, stockDespues: stockAntes + cant,
        requiereSerial: esSer,
        modoSerial: esSer ? modoSer : null,
        seriales: esSer && modoSer === 'individual' ? seriales : [],
        serialInicio: esSer && modoSer === 'rango' ? serialInicio : '',
        serialFin:    esSer && modoSer === 'rango' ? serialFin    : '',
        fecha: serverTimestamp(), registradoPor: session.uid, registradoPorNombre: session.displayName,
      });

      // Registrar un documento por serial para trazabilidad completa
      if (esSer) {
        let listaSeriales = [];
        if (modoSer === 'individual') {
          listaSeriales = seriales;
        } else {
          const nIni   = parseInt(serialInicio.replace(/[^0-9]/g,''), 10);
          const nFin   = parseInt(serialFin.replace(/[^0-9]/g,''), 10);
          const prefix = serialInicio.replace(/[0-9]+$/, '');
          const digits = String(nFin).length;
          for (let n = nIni; n <= nFin; n++) {
            listaSeriales.push(prefix + String(n).padStart(digits, '0'));
          }
        }
        const batch = listaSeriales.map(function(ser) {
          return addDoc(collection(db, 'kardex/seriales/items'), {
            sapCode: item.sapCode, axCode: item.axCode,
            itemId: item.id, itemNombre: safeStr(item?.name),
            serial: ser, estado: 'disponible',
            fechaEntrada: serverTimestamp(),
            registradoPor: session.uid,
          });
        });
        await Promise.all(batch);
      }

      ov.remove();
      showToast('Entrada de ' + cant + ' ' + safeStr(item?.unit,'') + ' registrada.', 'success');
      await showInventario(db, session);
    } catch(e) {
      errEl.textContent = 'Error al registrar.'; errEl.classList.remove('hidden');
      btn.disabled = false; btn.textContent = 'Registrar entrada'; console.error(e);
    }
  };
}

// ─────────────────────────────────────────
// FORM — NUEVA SALIDA (formato despacho)
// ─────────────────────────────────────────
async function showFormSalida(db, session) {
  const [snapI, snapU] = await Promise.all([
    getDocs(collection(db, 'kardex/inventario/items')),
    getDocs(query(collection(db, 'users'), where('active','==',true))),
  ]);

  const items = snapI.docs
    .map(d => normalizeItem({ id: d.id, ...d.data() }))
    .filter(i => esValido(i) && safeNum(i.stock) > 0)
    .sort((a,b) => safeStr(a.name).localeCompare(safeStr(b.name)));

  const usuarios = snapU.docs
    .map(d => d.data())
    .filter(u => ['campo','coordinadora'].includes(u.role))
    .sort((a,b) => safeStr(a.displayName).localeCompare(safeStr(b.displayName)));

  let sel   = [];
  let step  = 1; // 1 = encabezado, 2 = materiales
  let hdr   = {
    responsable: '', contratista: EMPRESAS_CONTRATISTAS[0]||'INNOVA',
    instalador: '', placaSel: '', placaOtro: '',
    fechaSol: today(), fechaEnt: today(),
  };
  let busq  = '';

  function tc(str) {
    return safeStr(str).toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  }
  function getRec() {
    try { return JSON.parse(sessionStorage.getItem('kardex_rec')||'[]'); } catch { return []; }
  }
  function addRec(id) {
    const p = getRec().filter(x=>x!==id);
    sessionStorage.setItem('kardex_rec', JSON.stringify([id,...p].slice(0,6)));
  }

  // ── Contenedor raíz ──
  const ov = document.createElement('div');
  ov.className = 'fixed inset-0 z-50';
  document.body.appendChild(ov);

  // ════════════════════════════════════════
  // PANTALLA 1 — ENCABEZADO
  // ════════════════════════════════════════
  function renderStep1() {
    ov.innerHTML = `
    <div class="flex flex-col h-full bg-white" style="max-height:100dvh">

      <!-- Barra superior -->
      <div class="flex items-center gap-3 px-4 py-3 border-b border-gray-100 shrink-0">
        <button id="s1-close" class="w-9 h-9 rounded-xl bg-gray-100 flex items-center justify-center shrink-0">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#374151" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
        <div class="flex-1">
          <p class="font-semibold text-gray-900">Nueva salida</p>
          <p class="text-xs text-gray-400">Paso 1 de 2 — Encabezado</p>
        </div>
        <!-- Indicador de pasos -->
        <div class="flex gap-1.5">
          <div class="w-6 h-1.5 rounded-full" style="background:#1B4F8A"></div>
          <div class="w-6 h-1.5 rounded-full bg-gray-200"></div>
        </div>
      </div>

      <!-- Campos -->
      <div class="flex-1 overflow-y-auto px-4 py-5 space-y-4">

        <!-- Usuario responsable — el más importante, va primero y grande -->
        <div>
          <label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Usuario responsable *</label>
          <div class="grid grid-cols-3 gap-2">
            ${USUARIOS_RESPONSABLES.map(u => `
            <button data-resp="${u}"
              class="py-3.5 rounded-2xl border-2 text-sm font-bold transition-all
                ${hdr.responsable===u
                  ? 'border-blue-500 text-white'
                  : 'border-gray-200 text-gray-700 bg-white'}"
              style="${hdr.responsable===u ? 'background:#1B4F8A' : ''}">
              ${u}
            </button>`).join('')}
          </div>
        </div>

        <!-- Instalador responsable -->
        <div>
          <label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Instalador responsable</label>
          <select id="s1-instalador"
            class="w-full border-2 border-gray-200 rounded-2xl px-4 py-3 text-sm bg-white focus:outline-none focus:border-blue-400 font-medium text-gray-800">
            <option value="">Seleccionar...</option>
            ${usuarios.map(u=>`<option value="${safeStr(u.displayName)}" ${hdr.instalador===safeStr(u.displayName)?'selected':''}>${safeStr(u.displayName)}</option>`).join('')}
          </select>
        </div>

        <!-- Empresa + Placa -->
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Empresa</label>
            <select id="s1-contratista"
              class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400 font-medium text-gray-800">
              ${EMPRESAS_CONTRATISTAS.map(e=>`<option value="${e}" ${hdr.contratista===e?'selected':''}>${e}</option>`).join('')}
              <option value="" ${hdr.contratista===''?'selected':''}>Otra</option>
            </select>
          </div>
          <div>
            <label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Placa</label>
            <select id="s1-placa"
              class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400 font-medium text-gray-800">
              <option value="">Seleccionar...</option>
              ${PLACAS.map(p=>`<option value="${p}" ${hdr.placaSel===p?'selected':''}>${p}</option>`).join('')}
              <option value="__otro__" ${hdr.placaSel==='__otro__'?'selected':''}>Otro</option>
            </select>
            <input id="s1-placa-otro" type="text" placeholder="Ej. P-123"
              value="${hdr.placaOtro}"
              class="${hdr.placaSel==='__otro__' ? '' : 'hidden'} w-full border-2 border-gray-200 rounded-2xl px-3 py-2.5 text-sm mt-2 font-mono focus:outline-none focus:border-blue-400"/>
          </div>
        </div>

        <!-- Fechas -->
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">F. solicitud</label>
            <input id="s1-fecha-sol" type="date" value="${hdr.fechaSol}"
              class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400"/>
          </div>
          <div>
            <label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">F. entrega</label>
            <input id="s1-fecha-ent" type="date" value="${hdr.fechaEnt}"
              class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400"/>
          </div>
        </div>

      </div>

      <!-- Botón siguiente -->
      <div class="px-4 py-4 border-t border-gray-100 shrink-0" style="padding-bottom:max(16px,env(safe-area-inset-bottom))">
        <div id="s1-err" class="hidden text-sm text-red-600 text-center mb-3"></div>
        <button id="s1-next"
          class="w-full font-bold rounded-2xl py-4 text-white flex items-center justify-center gap-2"
          style="background:#1B4F8A">
          Continuar — Agregar materiales
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
        </button>
      </div>
    </div>`;

    // Botones de usuario responsable
    ov.querySelectorAll('[data-resp]').forEach(b => {
      b.onclick = () => {
        hdr.responsable = b.dataset.resp;
        renderStep1();
      };
    });

    // Placa otro
    ov.querySelector('#s1-placa').addEventListener('change', e => {
      hdr.placaSel = e.target.value;
      const otro = ov.querySelector('#s1-placa-otro');
      if (e.target.value === '__otro__') { otro.classList.remove('hidden'); otro.focus(); }
      else { otro.classList.add('hidden'); hdr.placaOtro = ''; }
    });

    // Cerrar
    ov.querySelector('#s1-close').onclick = () => ov.remove();

    // Siguiente
    ov.querySelector('#s1-next').onclick = () => {
      // Guardar estado
      hdr.instalador  = ov.querySelector('#s1-instalador').value;
      hdr.contratista = ov.querySelector('#s1-contratista').value;
      hdr.placaSel    = ov.querySelector('#s1-placa').value;
      hdr.placaOtro   = ov.querySelector('#s1-placa-otro')?.value.trim() || '';
      hdr.fechaSol    = ov.querySelector('#s1-fecha-sol').value;
      hdr.fechaEnt    = ov.querySelector('#s1-fecha-ent').value;

      if (!hdr.responsable) {
        ov.querySelector('#s1-err').textContent = 'Selecciona el usuario responsable.';
        ov.querySelector('#s1-err').classList.remove('hidden');
        return;
      }
      step = 2;
      renderStep2();
    };
  }

  // ════════════════════════════════════════
  // PANTALLA 2 — MATERIALES
  // ════════════════════════════════════════
  function renderStep2() {
    const totalSel = sel.length;
    ov.innerHTML = `
    <div class="flex flex-col bg-gray-50" style="height:100dvh;max-height:100dvh">

      <!-- Barra superior -->
      <div class="flex items-center gap-3 px-4 py-3 border-b border-gray-100 bg-white shrink-0">
        <button id="s2-back" class="w-9 h-9 rounded-xl bg-gray-100 flex items-center justify-center shrink-0">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#374151" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/></svg>
        </button>
        <div class="flex-1 min-w-0">
          <p class="font-semibold text-gray-900">Materiales</p>
          <p class="text-xs text-gray-400 truncate">${hdr.responsable} · ${hdr.instalador||'Sin instalador'}</p>
        </div>
        <div class="flex gap-1.5">
          <div class="w-6 h-1.5 rounded-full bg-gray-200"></div>
          <div class="w-6 h-1.5 rounded-full" style="background:#1B4F8A"></div>
        </div>
      </div>

      <!-- Carrito (visible solo si hay items) -->
      ${totalSel > 0 ? `
      <div class="px-4 pt-3 pb-1 shrink-0 bg-white border-b border-gray-100">
        <div class="flex items-center justify-between mb-2">
          <p class="text-xs font-semibold text-gray-500 uppercase tracking-wider">En el despacho</p>
          <span class="text-xs font-bold px-2 py-0.5 rounded-full text-white" style="background:#1B4F8A">${totalSel}</span>
        </div>
        <div class="space-y-2 pb-1">
          ${sel.map((s,idx) => {
            const restante   = s.stock - s.cantidad;
            const restColor  = restante <= 0 ? '#C62828' : restante <= safeNum(items.find(i=>i.id===s.itemId)?.minStock||5) ? '#E65100' : '#166534';
            return `
            <div class="flex items-center gap-3 bg-gray-50 rounded-2xl px-3 py-2.5 border border-gray-200">
              <div class="flex-1 min-w-0">
                <p class="text-sm font-semibold text-gray-900 truncate leading-tight">${tc(s.name)}</p>
                <p class="text-xs font-medium mt-0.5" style="color:${restColor}">Restante: ${restante} ${safeStr(s.unit,'')}</p>
              </div>
              <!-- Controles cantidad -->
              <div class="flex items-center gap-1.5 shrink-0">
                <button data-dec="${idx}"
                  class="w-8 h-8 rounded-xl border border-gray-300 bg-white text-lg font-bold text-gray-600 flex items-center justify-center active:bg-gray-100">−</button>
                <span class="w-8 text-center text-base font-bold text-gray-900">${s.cantidad}</span>
                <button data-inc="${idx}"
                  class="w-8 h-8 rounded-xl border text-lg font-bold flex items-center justify-center active:opacity-70
                    ${s.cantidad>=s.stock ? 'border-gray-200 text-gray-300 bg-gray-50' : 'border-blue-300 text-blue-600 bg-blue-50'}"
                  ${s.cantidad>=s.stock?'disabled':''}>+</button>
              </div>
              <button data-del="${idx}"
                class="w-7 h-7 rounded-xl bg-red-50 flex items-center justify-center shrink-0">
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#EF4444" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              </button>
            </div>`;
          }).join('')}
        </div>
      </div>` : ''}

      <!-- Buscador + lista -->
      <div class="flex-1 overflow-y-auto px-4 pt-3 pb-2">
        <div class="relative mb-3">
          <svg class="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input id="s2-buscar" type="text" value="${busq}"
            placeholder="Buscar por nombre o SAP..."
            autocomplete="off"
            class="w-full bg-white border-2 border-gray-200 rounded-2xl pl-10 pr-4 py-3 text-sm focus:outline-none focus:border-blue-400 font-medium"/>
        </div>
        <div id="s2-lista" class="space-y-1.5 pb-2"></div>
      </div>

      <!-- Error -->
      <div id="s2-err" class="hidden mx-4 mb-1 text-sm text-red-600 text-center shrink-0"></div>

      <!-- Botón registrar sticky -->
      <div class="px-4 py-3 bg-white border-t border-gray-200 shrink-0"
        style="padding-bottom:max(16px,env(safe-area-inset-bottom))">
        <button id="s2-submit"
          class="w-full font-bold rounded-2xl py-4 text-sm text-white transition-all"
          style="background:${totalSel>0 ? '#1B4F8A' : '#D1D5DB'}"
          ${totalSel===0 ? 'disabled' : ''}>
          ${totalSel>0
            ? `Registrar salida · ${totalSel} material${totalSel!==1?'es':''}`
            : 'Agrega materiales para continuar'}
        </button>
      </div>
    </div>`;

    // Volver
    ov.querySelector('#s2-back').onclick = () => { step=1; renderStep1(); };

    // Carrito: dec / inc / del
    ov.querySelectorAll('[data-dec]').forEach(b => {
      b.onclick = () => { if (sel[+b.dataset.dec].cantidad > 1) { sel[+b.dataset.dec].cantidad--; renderStep2(); } };
    });
    ov.querySelectorAll('[data-inc]').forEach(b => {
      b.onclick = () => { if (sel[+b.dataset.inc].cantidad < sel[+b.dataset.inc].stock) { sel[+b.dataset.inc].cantidad++; renderStep2(); } };
    });
    ov.querySelectorAll('[data-del]').forEach(b => {
      b.onclick = () => { sel.splice(+b.dataset.del,1); renderStep2(); };
    });

    // Buscador
    ov.querySelector('#s2-buscar').addEventListener('input', e => { busq = e.target.value; renderLista(); });

    // Submit
    ov.querySelector('#s2-submit')?.addEventListener('click', handleSubmit);

    renderLista();
  }

  // ── Lista de materiales disponibles ──
  function renderLista() {
    const el = ov.querySelector('#s2-lista');
    if (!el) return;
    const selIds = new Set(sel.map(s=>s.itemId));
    const q = busq.trim().toLowerCase();

    let lista = q
      ? items.filter(i =>
          safeStr(i.name).toLowerCase().includes(q) ||
          safeStr(i.sapCode,'').includes(q) ||
          safeStr(i.axCode,'').includes(q))
      : (() => {
          const rec   = getRec().map(id=>items.find(i=>i.id===id)).filter(Boolean);
          const resto = items.filter(i=>!getRec().includes(i.id));
          return [...rec, ...resto];
        })();

    if (!lista.length) {
      el.innerHTML = '<p class="text-sm text-gray-400 text-center py-8">Sin resultados</p>';
      return;
    }

    el.innerHTML = lista.map(item => {
      const agregado   = selIds.has(item.id);
      const stockColor = safeNum(item.stock) <= safeNum(item.minStock||5) ? '#E65100' : '#374151';
      return `
        <button data-item="${item.id}" ${agregado?'disabled':''}
          class="w-full flex items-center justify-between gap-3 text-left px-4 py-3.5 rounded-2xl border-2 transition-all
            ${agregado
              ? 'border-green-200 bg-green-50 cursor-not-allowed'
              : 'border-gray-200 bg-white active:border-blue-400 active:bg-blue-50'}">
          <p class="text-sm font-semibold truncate ${agregado ? 'text-green-700' : 'text-gray-900'}">${tc(item.name)}</p>
          ${agregado
            ? `<span class="text-xs font-bold text-green-600 shrink-0">✓ Agregado</span>`
            : `<span class="text-sm font-bold shrink-0" style="color:${stockColor}">${item.stock}<span class="text-xs font-normal text-gray-400 ml-0.5">${safeStr(item.unit,'')}</span></span>`
          }
        </button>`;
    }).join('');

    el.querySelectorAll('[data-item]').forEach(btn => {
      btn.addEventListener('click', () => {
        const item = items.find(i=>i.id===btn.dataset.item);
        if (!item || sel.some(s=>s.itemId===item.id)) return;
        showCantidadModal(item);
      });
    });
  }

  // ── Advertencia stock usuario ──
  function buildStockWarning(item, hdr) {
    const resp = hdr.responsable;
    if (!resp || !window.__kardexStockUsuario) return '';
    const yaT = safeNum((window.__kardexStockUsuario[resp] || {})[item.id]);
    if (yaT <= 0) return '';
    const min        = safeNum(item.minStock || 5);
    const suficiente = yaT > min;
    const color = suficiente ? '#166534' : '#B45309';
    const bg    = suficiente ? '#F0FDF4'  : '#FFFBEB';
    const bdr   = suficiente ? '#BBF7D0'  : '#FDE68A';
    const icon  = suficiente ? '✅' : '⚠️';
    const unit  = safeStr(item.unit, '');
    const msg   = suficiente
      ? 'Este usuario ya tiene <strong>' + yaT + ' ' + unit + '</strong> — suficiente'
      : 'Este usuario ya tiene <strong>' + yaT + ' ' + unit + '</strong> — stock bajo';
    return '<div class="rounded-xl px-3 py-2 text-xs font-medium" style="background:' + bg + ';border:1px solid ' + bdr + ';color:' + color + '">' + icon + ' ' + msg + '</div>';
  }

  // ── Modal cantidad + seriales ──
  function showCantidadModal(item) {
    const esSello = item.sapCode === '354549';
    const esSer   = !!item.requiereSerial;
    let modo      = esSello ? 'rango' : 'individual';
    let cantVal   = 1;
    let serialesDisponibles = []; // cargados de Firestore

    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black/50 flex items-end z-50';
    document.body.appendChild(modal);
    modal.addEventListener('click', function(e) { if (e.target===modal) modal.remove(); });

    // Cargar seriales disponibles de Firestore si es serializable
    if (esSer) {
      getDocs(query(
        collection(db, 'kardex/seriales/items'),
        where('itemId', '==', item.id),
        where('estado', '==', 'disponible')
      )).then(function(snap) {
        serialesDisponibles = snap.docs.map(function(d) { return d.data().serial; }).sort();
        // Actualizar lista si ya está renderizada
        const lista = modal.querySelector('#mc-ser-lista');
        if (lista) renderSerialLista(lista, [], serialesDisponibles);
      }).catch(function(e) { console.error('Error cargando seriales:', e); });
    }

    function getSerialSeleccionados() {
      if (modo === 'rango') {
        const ini    = (modal.querySelector('#mc-ini')?.value || '').trim();
        const fin    = (modal.querySelector('#mc-fin')?.value || '').trim();
        const nIni   = parseInt(ini.replace(/[^0-9]/g,''), 10);
        const nFin   = parseInt(fin.replace(/[^0-9]/g,''), 10);
        const prefix = ini.replace(/[0-9]+$/, '');
        const digits = ini.replace(/[^0-9]/g,'').length;
        if (isNaN(nIni) || isNaN(nFin) || nFin < nIni) return [];
        const result = [];
        for (let n = nIni; n <= nFin; n++) result.push(prefix + String(n).padStart(digits, '0'));
        return result;
      } else {
        const raw = (modal.querySelector('#mc-sers')?.value || '').trim();
        return raw ? raw.split('\n').map(function(s){return s.trim();}).filter(Boolean) : [];
      }
    }

    function renderSerialLista(el, seleccionados, disponibles) {
      if (!disponibles.length) {
        el.innerHTML = '<p class="text-xs text-gray-400 text-center py-3">Sin seriales en bodega</p>';
        return;
      }
      const selSet  = new Set(seleccionados);
      const dispSet = new Set(disponibles);
      el.innerHTML = disponibles.map(function(s) {
        const enSel   = selSet.has(s);
        const bg      = enSel ? 'bg-green-50 border-green-300 text-green-700' : 'bg-gray-50 border-gray-200 text-gray-600';
        return '<span class="inline-block border rounded px-1.5 py-0.5 text-xs font-mono m-0.5 ' + bg + '">' + s + '</span>';
      }).join('') +
      (seleccionados.filter(function(s){ return !dispSet.has(s); }).length ?
        '<p class="text-xs text-red-500 mt-2">⚠ Algunos seriales no están disponibles en bodega</p>' : '');
    }

    function updateLista() {
      const lista = modal.querySelector('#mc-ser-lista');
      if (!lista) return;
      renderSerialLista(lista, getSerialSeleccionados(), serialesDisponibles);
      // También actualizar cantidad display
      const cant = getSerialSeleccionados().length;
      const disp = modal.querySelector('#mc-cant-display');
      if (disp) disp.textContent = cant || '—';
    }

    function calcCantFromSerial() {
      if (modo === 'rango') {
        const ini  = (modal.querySelector('#mc-ini')?.value || '').trim();
        const fin  = (modal.querySelector('#mc-fin')?.value || '').trim();
        const nIni = parseInt(ini.replace(/[^0-9]/g,''), 10);
        const nFin = parseInt(fin.replace(/[^0-9]/g,''), 10);
        return (!isNaN(nIni) && !isNaN(nFin) && nFin >= nIni) ? (nFin - nIni + 1) : 0;
      } else {
        const raw = (modal.querySelector('#mc-sers')?.value || '').trim();
        return raw ? raw.split('\n').map(function(s){return s.trim();}).filter(Boolean).length : 0;
      }
    }

    function updateCantDisplay() {
      const disp = modal.querySelector('#mc-cant-display');
      if (disp) disp.textContent = calcCantFromSerial() || '—';
    }

    function render() {
      const sheet = document.createElement('div');
      sheet.className = 'bg-white w-full rounded-t-3xl px-5 pt-5';
      sheet.style.cssText = 'padding-bottom:max(32px,env(safe-area-inset-bottom));max-height:90dvh;overflow-y:auto';

      const handle = document.createElement('div');
      handle.className = 'w-10 h-1 bg-gray-200 rounded-full mx-auto mb-4';
      sheet.appendChild(handle);

      const title = document.createElement('div');
      title.className = 'mb-4';
      title.innerHTML = '<p class="font-bold text-gray-900 text-lg leading-tight">' + tc(item.name) + '</p>' +
        '<p class="text-sm text-gray-400 mt-0.5">' + item.stock + ' ' + safeStr(item.unit,'') + ' en bodega</p>';
      sheet.appendChild(title);

      const warn = document.createElement('div');
      warn.innerHTML = buildStockWarning(item, hdr);
      if (warn.innerHTML) sheet.appendChild(warn);

      // Cantidad — editable solo si no es serializable
      const cantRow = document.createElement('div');
      cantRow.className = 'mb-4';
      if (esSer) {
        cantRow.innerHTML =
          '<div class="bg-blue-50 border border-blue-100 rounded-xl p-3 text-center">' +
            '<p class="text-xs text-blue-500 mb-0.5">Cantidad calculada del serial</p>' +
            '<p id="mc-cant-display" class="text-3xl font-black text-blue-700">—</p>' +
            '<input id="mc-cant" type="hidden" value="0"/>' +
          '</div>';
      } else {
        cantRow.className = 'flex items-center justify-between gap-4 mb-4';
        cantRow.innerHTML =
          '<button id="mc-dec" class="w-16 h-16 rounded-2xl border-2 border-gray-200 bg-gray-50 text-3xl font-bold text-gray-400 flex items-center justify-center active:bg-gray-200">−</button>' +
          '<div class="flex-1 text-center">' +
            '<input id="mc-cant" type="number" min="1" max="' + item.stock + '" value="' + cantVal + '" class="w-full text-center text-5xl font-black text-gray-900 bg-transparent border-none focus:outline-none leading-none py-2"/>' +
            '<p class="text-sm text-gray-400 -mt-1">' + safeStr(item.unit,'') + '</p>' +
          '</div>' +
          '<button id="mc-inc" class="w-16 h-16 rounded-2xl border-2 border-blue-300 bg-blue-50 text-3xl font-bold text-blue-600 flex items-center justify-center active:bg-blue-100">+</button>';
      }
      sheet.appendChild(cantRow);

      // Seriales
      if (esSer) {
        const serDiv = document.createElement('div');
        serDiv.className = 'border-t border-gray-100 pt-4 space-y-3';

        const serTitle = document.createElement('p');
        serTitle.className = 'text-sm font-semibold text-gray-700';
        serTitle.textContent = 'Seriales / Sellos';
        serDiv.appendChild(serTitle);

        if (!esSello) {
          const modoRow = document.createElement('div');
          modoRow.className = 'flex gap-2';
          modoRow.innerHTML =
            '<button id="mc-modo-ind" class="flex-1 py-2 rounded-xl text-sm font-semibold border-2 transition-all ' +
              (modo==='individual' ? 'border-blue-500 bg-blue-50 text-blue-700' : 'border-gray-200 bg-white text-gray-500') + '">Individual</button>' +
            '<button id="mc-modo-rng" class="flex-1 py-2 rounded-xl text-sm font-semibold border-2 transition-all ' +
              (modo==='rango' ? 'border-blue-500 bg-blue-50 text-blue-700' : 'border-gray-200 bg-white text-gray-500') + '">Rango</button>';
          serDiv.appendChild(modoRow);
        }

        if (modo === 'rango') {
          const rangoRow = document.createElement('div');
          rangoRow.className = 'flex gap-2';
          rangoRow.innerHTML =
            '<div class="flex-1"><p class="text-xs text-gray-400 mb-1">Serial inicio</p>' +
              '<input id="mc-ini" type="text" placeholder="Ej: 12345001" class="w-full border-2 border-gray-200 rounded-xl px-3 py-2.5 text-sm font-mono focus:outline-none focus:border-blue-400"/></div>' +
            '<div class="flex-1"><p class="text-xs text-gray-400 mb-1">Serial fin</p>' +
              '<input id="mc-fin" type="text" placeholder="Ej: 12345030" class="w-full border-2 border-gray-200 rounded-xl px-3 py-2.5 text-sm font-mono focus:outline-none focus:border-blue-400"/></div>';
          serDiv.appendChild(rangoRow);
        } else {
          const indDiv = document.createElement('div');
          indDiv.innerHTML =
            '<p class="text-xs text-gray-400 mb-1">Seriales (uno por línea)</p>' +
            '<textarea id="mc-sers" rows="4" placeholder="12345001&#10;12345002&#10;12345003" class="w-full border-2 border-gray-200 rounded-xl px-3 py-2 text-sm font-mono focus:outline-none focus:border-blue-400 resize-none"></textarea>';
          serDiv.appendChild(indDiv);
        }
        // Lista de seriales disponibles
        const listaDiv = document.createElement('div');
        listaDiv.className = 'mt-3';
        const listaTitle = document.createElement('p');
        listaTitle.className = 'text-xs text-gray-500 font-medium mb-1';
        listaTitle.textContent = 'Disponibles en bodega:';
        const listaEl = document.createElement('div');
        listaEl.id = 'mc-ser-lista';
        listaEl.className = 'max-h-24 overflow-y-auto';
        listaEl.innerHTML = '<p class="text-xs text-gray-400">Cargando...</p>';
        listaDiv.appendChild(listaTitle);
        listaDiv.appendChild(listaEl);
        serDiv.appendChild(listaDiv);

        sheet.appendChild(serDiv);
      }

      const errEl = document.createElement('div');
      errEl.id = 'mc-err';
      errEl.className = 'hidden text-sm text-red-500 text-center mt-3';
      sheet.appendChild(errEl);

      const addBtn = document.createElement('button');
      addBtn.id = 'mc-add';
      addBtn.className = 'w-full text-white font-bold rounded-2xl py-4 text-base active:opacity-90 mt-4';
      addBtn.style.background = '#1B4F8A';
      addBtn.textContent = 'Agregar al despacho';
      sheet.appendChild(addBtn);

      modal.innerHTML = '';
      modal.appendChild(sheet);

      // Renderizar lista inicial si ya cargó
      if (esSer && serialesDisponibles.length) {
        const lista = modal.querySelector('#mc-ser-lista');
        if (lista) renderSerialLista(lista, [], serialesDisponibles);
      }

      if (!esSer) {
        const cantEl = modal.querySelector('#mc-cant');
        setTimeout(function() { cantEl.focus(); cantEl.select(); }, 80);
        modal.querySelector('#mc-dec').onclick = function() { cantVal = Math.max(1, safeNum(cantEl.value)-1); cantEl.value = cantVal; };
        modal.querySelector('#mc-inc').onclick = function() { cantVal = Math.min(item.stock, safeNum(cantEl.value)+1); cantEl.value = cantVal; };
        cantEl.oninput = function() { cantVal = safeNum(cantEl.value); };
      } else {
        // Auto-calcular cantidad y actualizar lista desde serial
        setTimeout(function() {
          modal.querySelector('#mc-ini')?.addEventListener('input', updateLista);
          modal.querySelector('#mc-fin')?.addEventListener('input', updateLista);
          modal.querySelector('#mc-sers')?.addEventListener('input', updateLista);
        }, 50);
      }

      if (!esSello) {
        modal.querySelector('#mc-modo-ind')?.addEventListener('click', function() { modo='individual'; render(); });
        modal.querySelector('#mc-modo-rng')?.addEventListener('click', function() { modo='rango'; render(); });
      }

      addBtn.addEventListener('click', function() {
        let cant = 0;
        let seriales=[], serialInicio='', serialFin='';

        if (esSer) {
          if (modo === 'rango') {
            serialInicio = (modal.querySelector('#mc-ini')?.value || '').trim();
            serialFin    = (modal.querySelector('#mc-fin')?.value || '').trim();
            if (!serialInicio || !serialFin) { errEl.textContent='Ingresa serial inicio y fin.'; errEl.classList.remove('hidden'); return; }
            const nIni = parseInt(serialInicio.replace(/[^0-9]/g,''), 10);
            const nFin = parseInt(serialFin.replace(/[^0-9]/g,''), 10);
            cant = (!isNaN(nIni) && !isNaN(nFin) && nFin >= nIni) ? (nFin - nIni + 1) : 0;
            if (cant <= 0) { errEl.textContent='Rango inválido.'; errEl.classList.remove('hidden'); return; }
          } else {
            const raw = (modal.querySelector('#mc-sers')?.value || '').trim();
            seriales = raw ? raw.split('\n').map(function(s){return s.trim();}).filter(Boolean) : [];
            if (!seriales.length) { errEl.textContent='Ingresa al menos un serial.'; errEl.classList.remove('hidden'); return; }
            cant = seriales.length;
          }
          if (cant > item.stock) { errEl.textContent='Supera el stock disponible ('+item.stock+').'; errEl.classList.remove('hidden'); return; }
          // Validar que todos los seriales existen en bodega
          if (serialesDisponibles.length > 0) {
            const dispSet = new Set(serialesDisponibles);
            const noDisp  = getSerialSeleccionados().filter(function(s){ return !dispSet.has(s); });
            if (noDisp.length) {
              errEl.textContent = 'Seriales no disponibles: ' + noDisp.slice(0,3).join(', ') + (noDisp.length>3?' y '+(noDisp.length-3)+' más':'');
              errEl.classList.remove('hidden'); return;
            }
          }
        } else {
          cant = safeNum(modal.querySelector('#mc-cant')?.value);
          if (cant <= 0)         { errEl.textContent='Ingresa una cantidad mayor a 0.'; errEl.classList.remove('hidden'); return; }
          if (cant > item.stock) { errEl.textContent='Máximo: '+item.stock+' '+safeStr(item.unit,''); errEl.classList.remove('hidden'); return; }
        }

        sel.push({
          itemId:item.id, name:item.name, unit:item.unit,
          sapCode:item.sapCode, axCode:item.axCode, stock:item.stock,
          cantidad:cant, requiereSerial:esSer,
          modoSerial: esSer ? modo : 'individual',
          seriales:seriales, serialInicio:serialInicio, serialFin:serialFin,
        });
        addRec(item.id);
        modal.remove();
        renderStep2();
      });

      if (!esSer) {
        modal.querySelector('#mc-cant')?.addEventListener('keydown', function(e) { if(e.key==='Enter') addBtn.click(); });
      }
    }

    render();
  }

  // ── Submit ──
  async function handleSubmit() {
    const errEl = ov.querySelector('#s2-err');
    const btn   = ov.querySelector('#s2-submit');
    if (!sel.length) return;
    errEl?.classList.add('hidden');
    btn.disabled = true;
    btn.textContent = 'Registrando...';
    const placa = hdr.placaSel==='__otro__' ? hdr.placaOtro : hdr.placaSel;
    try {
      const ref = await addDoc(collection(db, 'kardex/movimientos/salidas'), {
        usuarioResponsableUid: hdr.responsable,
        usuarioResponsable:    hdr.responsable,
        empresaContratista:    hdr.contratista,
        instaladorResponsable: hdr.instalador,
        placaVehiculo:         placa,
        fechaSolicitud: hdr.fechaSol,
        fechaEntrega:   hdr.fechaEnt,
        entregadoPor:    session.displayName,
        entregadoPorUid: session.uid,
        items: sel.map(s=>({
          itemId:s.itemId, sapCode:s.sapCode, axCode:s.axCode,
          nombre:s.name, unit:s.unit, cantidad:s.cantidad,
          requiereSerial:s.requiereSerial,
          modoSerial:   s.requiereSerial ? s.modoSerial    : null,
          seriales:     s.requiereSerial && s.modoSerial==='individual' ? s.seriales : [],
          serialInicio: s.requiereSerial && s.modoSerial==='rango' ? s.serialInicio : '',
          serialFin:    s.requiereSerial && s.modoSerial==='rango' ? s.serialFin    : '',
        })),
        fecha: serverTimestamp(),
      });
      for (const s of sel) {
        await updateDoc(doc(db,'kardex/inventario/items',s.itemId),{stock:increment(-s.cantidad)});
      }

      // Marcar seriales como despachados en trazabilidad
      for (const s of sel) {
        if (!s.requiereSerial) continue;
        let listaSeriales = s.seriales || [];
        if (s.modoSerial === 'rango' && s.serialInicio) {
          const nIni   = parseInt(s.serialInicio.replace(/[^0-9]/g,''), 10);
          const nFin   = parseInt((s.serialFin||'').replace(/[^0-9]/g,''), 10);
          const prefix = s.serialInicio.replace(/[0-9]+$/, '');
          const digits = s.serialInicio.replace(/[^0-9]/g,'').length;
          listaSeriales = [];
          for (let n = nIni; n <= nFin; n++) {
            listaSeriales.push(prefix + String(n).padStart(digits, '0'));
          }
        }
        if (!listaSeriales.length) continue;
        const snapSer = await getDocs(query(
          collection(db, 'kardex/seriales/items'),
          where('itemId', '==', s.itemId),
          where('estado', '==', 'disponible')
        ));
        const serSet = new Set(listaSeriales);
        const updates = snapSer.docs
          .filter(function(d){ return serSet.has(d.data().serial); })
          .map(function(d){
            return updateDoc(d.ref, {
              estado: 'despachado',
              salidaId: ref.id,
              fechaSalida: serverTimestamp(),
              usuarioDespacho: hdr.responsable,
            });
          });
        await Promise.all(updates);
      }

      ov.remove();
      showToast('Salida registrada correctamente.','success');
      showMemo({
        id:ref.id, usuarioResponsable:hdr.responsable,
        empresaContratista:hdr.contratista, instaladorResponsable:hdr.instalador,
        entregadoPor:session.displayName, placaVehiculo:placa,
        fechaSolicitud:hdr.fechaSol, fechaEntrega:hdr.fechaEnt, items:sel,
      });
      await showDashboard(db, session);
    } catch(e) {
      if(errEl){ errEl.textContent='Error al registrar. Intenta de nuevo.'; errEl.classList.remove('hidden'); }
      btn.disabled=false;
      btn.textContent=`Registrar salida · ${sel.length} material${sel.length!==1?'es':''}`;
      console.error(e);
    }
  }

  // Init
  renderStep1();
}

// ─────────────────────────────────────────
// HISTORIAL
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// DEVOLUCIÓN DIRECTA (admin/coordinadora)
// ─────────────────────────────────────────
async function showFormDevolucion(db, session, salida) {
  // Build items list with serial support
  const items = (salida.items || []);
  let selDev = {}; // itemId -> { cantidad, seriales[], serialInicio, serialFin, modoSerial }
  items.forEach(function(i) { selDev[i.itemId] = { cantidad:0, seriales:[], serialInicio:'', serialFin:'', modoSerial:'individual', requiereSerial:!!i.requiereSerial, nombre:i.nombre||i.name, unit:i.unit||i.unidad, sapCode:i.sapCode, cantidadOriginal:i.cantidad }; });

  function buildItemsHTML() {
    return items.map(function(it) {
      const sd = selDev[it.itemId];
      const esSello = it.sapCode === '354549';
      const esSer   = !!it.requiereSerial;
      return '<div class="border border-gray-200 rounded-xl p-3 space-y-2" id="dev-item-' + it.itemId + '">' +
        '<div class="flex items-center justify-between gap-2">' +
          '<p class="text-sm font-medium text-gray-900 flex-1 leading-tight">' + safeStr(it.nombre||it.name) + '</p>' +
          '<label class="flex items-center gap-1.5 shrink-0">' +
            '<input type="checkbox" class="dev-chk w-4 h-4 rounded" style="accent-color:#1B4F8A" data-iid="' + it.itemId + '" ' + (sd.cantidad > 0 ? 'checked' : '') + '/>' +
            '<span class="text-xs text-gray-500">Devolver</span>' +
          '</label>' +
        '</div>' +
        '<div class="dev-campos-' + it.itemId + (sd.cantidad > 0 ? '' : ' hidden') + '">' +
          (!esSer ?
            '<div>' +
              '<label class="text-xs text-gray-500">Cantidad a devolver (máx ' + it.cantidad + ')</label>' +
              '<input type="number" min="1" max="' + it.cantidad + '" value="' + (sd.cantidad||1) + '" class="dev-cant w-full border border-gray-200 rounded-lg px-3 py-2 text-sm text-center font-semibold mt-1 focus:outline-none focus:ring-2 focus:ring-blue-400" data-iid="' + it.itemId + '"/>' +
            '</div>'
          :
            '<div class="space-y-2">' +
              (!esSello ?
                '<div class="flex gap-2">' +
                  '<button class="dev-modo-ind flex-1 py-1.5 rounded-lg text-xs font-semibold border-2 ' + (sd.modoSerial==='individual' ? 'border-blue-500 bg-blue-50 text-blue-700' : 'border-gray-200 text-gray-500') + '" data-iid="' + it.itemId + '">Individual</button>' +
                  '<button class="dev-modo-rng flex-1 py-1.5 rounded-lg text-xs font-semibold border-2 ' + (sd.modoSerial==='rango' ? 'border-blue-500 bg-blue-50 text-blue-700' : 'border-gray-200 text-gray-500') + '" data-iid="' + it.itemId + '">Rango</button>' +
                '</div>'
              : '') +
              (sd.modoSerial === 'rango' ?
                '<div class="flex gap-2">' +
                  '<div class="flex-1"><p class="text-xs text-gray-400 mb-1">Serial inicio</p><input type="text" class="dev-ini w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs font-mono focus:outline-none focus:ring-2 focus:ring-blue-400" data-iid="' + it.itemId + '" value="' + (sd.serialInicio||'') + '"/></div>' +
                  '<div class="flex-1"><p class="text-xs text-gray-400 mb-1">Serial fin</p><input type="text" class="dev-fin w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs font-mono focus:outline-none focus:ring-2 focus:ring-blue-400" data-iid="' + it.itemId + '" value="' + (sd.serialFin||'') + '"/></div>' +
                '</div>'
              :
                '<div><p class="text-xs text-gray-400 mb-1">Seriales (uno por línea)</p><textarea class="dev-sers w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs font-mono focus:outline-none focus:ring-2 focus:ring-blue-400 resize-none" rows="3" data-iid="' + it.itemId + '">' + (sd.seriales||[]).join('\n') + '</textarea></div>'
              ) +
              '<div class="bg-blue-50 rounded-lg px-3 py-2 text-center">' +
                '<p class="text-xs text-blue-500">Cantidad calculada</p>' +
                '<p class="dev-cant-display text-lg font-black text-blue-700" data-iid="' + it.itemId + '">—</p>' +
              '</div>' +
            '</div>'
          ) +
        '</div>' +
      '</div>';
    }).join('');
  }

  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div>' +
        '<h2 class="font-semibold text-gray-900">Registrar devolución</h2>' +
        '<p class="text-xs text-gray-400 mt-0.5">' + safeStr(salida.usuarioResponsable) + ' · ' + fmtDate(salida.fecha) + '</p>' +
      '</div>' +
      '<button id="dv-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-3">' +
      '<div id="dv-items">' + buildItemsHTML() + '</div>' +
      '<div>' +
        '<label class="text-xs text-gray-500 font-medium">Motivo (opcional)</label>' +
        '<textarea id="dv-motivo" rows="2" placeholder="Ej. Material sobrante, equipo defectuoso..." class="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm mt-1 focus:outline-none focus:ring-2 focus:ring-blue-400 resize-none"></textarea>' +
      '</div>' +
      '<div id="dv-err" class="hidden text-sm text-red-500 bg-red-50 rounded-xl px-3 py-2"></div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex gap-3 shrink-0">' +
      '<button id="dv-cancel" class="flex-1 border border-gray-200 text-gray-600 font-medium rounded-xl py-3 text-sm">Cancelar</button>' +
      '<button id="dv-submit" class="flex-1 text-white font-medium rounded-xl py-3 text-sm" style="background:#B45309">Registrar devolución</button>' +
    '</div>'
  );

  ov.querySelector('#dv-close').onclick = ov.querySelector('#dv-cancel').onclick = () => ov.remove();

  function calcCant(iid) {
    const sd = selDev[iid];
    if (!sd.requiereSerial) return sd.cantidad;
    if (sd.modoSerial === 'rango') {
      const nI = parseInt((sd.serialInicio||'').replace(/[^0-9]/g,''), 10);
      const nF = parseInt((sd.serialFin||'').replace(/[^0-9]/g,''), 10);
      return (!isNaN(nI) && !isNaN(nF) && nF >= nI) ? (nF - nI + 1) : 0;
    }
    return (sd.seriales||[]).length;
  }

  function updateDisplay(iid) {
    const disp = ov.querySelector('.dev-cant-display[data-iid="' + iid + '"]');
    if (disp) disp.textContent = calcCant(iid) || '—';
  }

  function wireItem(iid) {
    const sd = selDev[iid];
    // Checkbox toggle
    const chk = ov.querySelector('.dev-chk[data-iid="' + iid + '"]');
    const campos = ov.querySelector('.dev-campos-' + iid);
    if (chk) chk.addEventListener('change', function() {
      if (this.checked) { campos.classList.remove('hidden'); if (!sd.requiereSerial) sd.cantidad = 1; }
      else { campos.classList.add('hidden'); sd.cantidad = 0; }
    });
    // Cantidad (no serializable)
    const cantEl = ov.querySelector('.dev-cant[data-iid="' + iid + '"]');
    if (cantEl) cantEl.addEventListener('input', function() { sd.cantidad = safeNum(this.value); });
    // Modo buttons
    const indBtn = ov.querySelector('.dev-modo-ind[data-iid="' + iid + '"]');
    const rngBtn = ov.querySelector('.dev-modo-rng[data-iid="' + iid + '"]');
    if (indBtn) indBtn.addEventListener('click', function() {
      sd.modoSerial = 'individual';
      ov.querySelector('#dv-items').innerHTML = buildItemsHTML();
      items.forEach(function(i){ wireItem(i.itemId); });
    });
    if (rngBtn) rngBtn.addEventListener('click', function() {
      sd.modoSerial = 'rango';
      ov.querySelector('#dv-items').innerHTML = buildItemsHTML();
      items.forEach(function(i){ wireItem(i.itemId); });
    });
    // Serial inputs
    const iniEl = ov.querySelector('.dev-ini[data-iid="' + iid + '"]');
    const finEl = ov.querySelector('.dev-fin[data-iid="' + iid + '"]');
    const sersEl= ov.querySelector('.dev-sers[data-iid="' + iid + '"]');
    if (iniEl) iniEl.addEventListener('input', function() { sd.serialInicio = this.value.trim(); updateDisplay(iid); });
    if (finEl) finEl.addEventListener('input', function() { sd.serialFin = this.value.trim(); updateDisplay(iid); });
    if (sersEl) sersEl.addEventListener('input', function() {
      sd.seriales = this.value.trim().split('\n').map(function(s){return s.trim();}).filter(Boolean);
      updateDisplay(iid);
    });
  }

  items.forEach(function(i){ wireItem(i.itemId); });

  ov.querySelector('#dv-submit').onclick = async function() {
    const errEl = ov.querySelector('#dv-err');
    const btn   = ov.querySelector('#dv-submit');
    const motivo = ov.querySelector('#dv-motivo').value.trim();
    errEl.classList.add('hidden');

    // Collect items to return
    const devItems = [];
    for (const it of items) {
      const sd = selDev[it.itemId];
      const chk = ov.querySelector('.dev-chk[data-iid="' + it.itemId + '"]');
      if (!chk || !chk.checked) continue;

      let cant = 0, seriales = [], serialInicio = '', serialFin = '';
      if (sd.requiereSerial) {
        if (sd.modoSerial === 'rango') {
          serialInicio = sd.serialInicio;
          serialFin    = sd.serialFin;
          const nI = parseInt(serialInicio.replace(/[^0-9]/g,''), 10);
          const nF = parseInt(serialFin.replace(/[^0-9]/g,''), 10);
          cant = (!isNaN(nI) && !isNaN(nF) && nF >= nI) ? (nF - nI + 1) : 0;
          if (!serialInicio || !serialFin || cant <= 0) { errEl.textContent = 'Ingresa rango válido para ' + safeStr(it.nombre||it.name); errEl.classList.remove('hidden'); return; }
        } else {
          seriales = sd.seriales || [];
          cant = seriales.length;
          if (!cant) { errEl.textContent = 'Ingresa seriales para ' + safeStr(it.nombre||it.name); errEl.classList.remove('hidden'); return; }
        }
        if (cant > it.cantidad) { errEl.textContent = 'No puedes devolver más de ' + it.cantidad + ' de ' + safeStr(it.nombre||it.name); errEl.classList.remove('hidden'); return; }
      } else {
        cant = sd.cantidad;
        if (!cant || cant <= 0) { errEl.textContent = 'Ingresa cantidad válida para ' + safeStr(it.nombre||it.name); errEl.classList.remove('hidden'); return; }
        if (cant > it.cantidad) { errEl.textContent = 'No puedes devolver más de ' + it.cantidad + ' de ' + safeStr(it.nombre||it.name); errEl.classList.remove('hidden'); return; }
      }

      devItems.push({ itemId:it.itemId, nombre:it.nombre||it.name, unit:it.unit||it.unidad, sapCode:it.sapCode, cantidad:cant, requiereSerial:!!it.requiereSerial, modoSerial:sd.modoSerial, seriales, serialInicio, serialFin });
    }

    if (!devItems.length) { errEl.textContent = 'Selecciona al menos un material a devolver.'; errEl.classList.remove('hidden'); return; }

    btn.disabled = true; btn.textContent = 'Registrando...';
    try {
      // 1. Update stock
      for (const di of devItems) {
        await updateDoc(doc(db, 'kardex/inventario/items', di.itemId), { stock: increment(di.cantidad) });
      }

      // 2. Reactivate serials
      for (const di of devItems) {
        if (!di.requiereSerial) continue;
        let lista = di.seriales;
        if (di.modoSerial === 'rango') {
          const nI = parseInt(di.serialInicio.replace(/[^0-9]/g,''), 10);
          const nF = parseInt(di.serialFin.replace(/[^0-9]/g,''), 10);
          const prefix = di.serialInicio.replace(/[0-9]+$/, '');
          const digits = di.serialInicio.replace(/[^0-9]/g,'').length;
          lista = [];
          for (let n = nI; n <= nF; n++) lista.push(prefix + String(n).padStart(digits, '0'));
        }
        if (!lista.length) continue;
        const snap = await getDocs(query(
          collection(db, 'kardex/seriales/items'),
          where('itemId', '==', di.itemId),
          where('estado', '==', 'despachado')
        ));
        const serSet = new Set(lista);
        const updates = snap.docs
          .filter(function(d){ return serSet.has(d.data().serial); })
          .map(function(d){ return updateDoc(d.ref, { estado:'disponible', salidaId:null, fechaSalida:null, usuarioDespacho:null }); });
        await Promise.all(updates);
      }

      // 3. Register movement
      await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
        tipo: 'devolucion',
        salidaOrigen: salida.id,
        usuarioResponsable: safeStr(salida.usuarioResponsable),
        items: devItems,
        motivo: motivo || 'Sin motivo',
        registradoPor: session.uid,
        registradoPorNombre: session.displayName,
        fecha: serverTimestamp(),
      });

      ov.remove();
      showToast('Devolución registrada correctamente.', 'success');
      await showHistorial(db, session);
    } catch(e) {
      errEl.textContent = 'Error al registrar. Intenta de nuevo.';
      errEl.classList.remove('hidden');
      btn.disabled = false; btn.textContent = 'Registrar devolución';
      console.error(e);
    }
  };
}

async function showHistorial(db, session) {
  setTab('historial');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  try {
    const [snapSalidas, snapDev] = await Promise.all([
      getDocs(query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'))),
      getDocs(query(collection(db, 'kardex/movimientos/ajustes'), where('tipo','==','devolucion'))),
    ]);
    const salidas = snapSalidas.docs.map(d => ({ id: d.id, _tipo:'salida', ...d.data() }));
    const devs    = snapDev.docs.map(d => ({ id: d.id, _tipo:'devolucion', ...d.data() }));

    // Merge and sort by fecha desc
    const todos = [...salidas, ...devs].sort((a, b) => {
      const ta = a.fecha?.toMillis?.() || 0;
      const tb = b.fecha?.toMillis?.() || 0;
      return tb - ta;
    });

    function rowSalida(s) {
      return '<div class="px-4 py-3">' +
        '<div class="flex items-start justify-between gap-2">' +
          '<div class="min-w-0">' +
            '<div class="flex items-center gap-1.5">' +
              '<span class="text-xs font-semibold px-1.5 py-0.5 rounded" style="background:#DCFCE7;color:#166534">↑ Salida</span>' +
              '<p class="text-sm font-semibold text-gray-900">' + safeStr(s.usuarioResponsable||s.tecnicoNombre) + '</p>' +
            '</div>' +
            '<p class="text-xs text-gray-400 mt-0.5">' + fmtDate(s.fecha) + ' · ' + safeStr(s.entregadoPor||s.registradoPorNombre) + '</p>' +
            (s.empresaContratista ? '<p class="text-xs text-gray-400">' + s.empresaContratista + '</p>' : '') +
          '</div>' +
          '<div class="flex gap-1.5 shrink-0">' +
            '<button data-smemo="' + s.id + '" class="text-xs font-medium px-2 py-1 rounded-lg" style="color:#1B4F8A;background:#EFF6FF">Memo</button>' +
            '<button data-sdev="' + s.id + '" class="text-xs font-medium px-2 py-1 rounded-lg" style="color:#B45309;background:#FEF3C7">Devolver</button>' +
          '</div>' +
        '</div>' +
        '<div class="flex flex-wrap gap-1 mt-2">' +
          (s.items||[]).map(function(i) {
            return '<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">' + safeNum(i.cantidad) + ' ' + safeStr(i.unit||i.unidad,'') + ' ' + safeStr(i.nombre||i.name) + '</span>';
          }).join('') +
        '</div>' +
      '</div>';
    }

    function rowDevolucion(d) {
      return '<div class="px-4 py-3">' +
        '<div class="flex items-start justify-between gap-2">' +
          '<div class="min-w-0">' +
            '<div class="flex items-center gap-1.5">' +
              '<span class="text-xs font-semibold px-1.5 py-0.5 rounded" style="background:#FEE2E2;color:#C62828">↩ Devolución</span>' +
              '<p class="text-sm font-semibold text-gray-900">' + safeStr(d.usuarioResponsable) + '</p>' +
            '</div>' +
            '<p class="text-xs text-gray-400 mt-0.5">' + fmtDate(d.fecha) + ' · ' + safeStr(d.registradoPorNombre) + '</p>' +
            (d.motivo ? '<p class="text-xs text-gray-500 mt-0.5 italic">' + d.motivo + '</p>' : '') +
          '</div>' +
        '</div>' +
        '<div class="flex flex-wrap gap-1 mt-2">' +
          (d.items||[]).map(function(i) {
            return '<span class="text-xs px-2 py-0.5 rounded-full text-red-700" style="background:#FEE2E2">' + safeNum(i.cantidad) + ' ' + safeStr(i.unit,'') + ' ' + safeStr(i.nombre) + '</span>';
          }).join('') +
        '</div>' +
      '</div>';
    }

    content.innerHTML =
      '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
        '<div class="px-4 py-3 border-b border-gray-100">' +
          '<p class="font-semibold text-sm text-gray-900">Historial de movimientos</p>' +
          '<p class="text-xs text-gray-400 mt-0.5">' + todos.length + ' registro(s)</p>' +
        '</div>' +
        (todos.length === 0
          ? '<div class="py-12 text-center text-sm text-gray-400">Sin movimientos registrados</div>'
          : '<div class="divide-y divide-gray-100">' +
              todos.map(function(m) { return m._tipo === 'salida' ? rowSalida(m) : rowDevolucion(m); }).join('') +
            '</div>'
        ) +
      '</div>';

    const salidaMap = {};
    salidas.forEach(s => { salidaMap[s.id] = s; });
    content.querySelectorAll('[data-smemo]').forEach(b => {
      b.onclick = () => {
        const s = salidaMap[b.dataset.smemo];
        if (s) showMemo({ ...s, fecha: s.fecha?.toDate?.() || new Date() });
      };
    });
    content.querySelectorAll('[data-sdev]').forEach(b => {
      b.onclick = () => {
        const s = salidaMap[b.dataset.sdev];
        if (s) showFormDevolucion(db, session, s);
      };
    });
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// MEMO — Formato documento físico DELSUR
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// MAPEO: datos internos → formato oficial
// Solo los campos del documento corporativo.
// ─────────────────────────────────────────
function mapearParaMemo(s) {
  const hoy = new Date().toLocaleDateString('es-SV', { year:'numeric', month:'long', day:'numeric' });
  return {
    // Campos del encabezado oficial — en el mismo orden del documento físico
    USUARIO_RESPONSABLE:    safeStr(s.usuarioResponsable || s.tecnicoNombre, ''),
    EMPRESA_CONTRATISTA:    safeStr(s.empresaContratista, ''),
    INSTALADOR_RESPONSABLE: safeStr(s.instaladorResponsable, ''),
    ENTREGADO_POR:          safeStr(s.entregadoPor || s.registradoPorNombre, ''),
    PLACA_VEHICULO:         safeStr(s.placaVehiculo, ''),
    FECHA_SOLICITUD:        safeStr(s.fechaSolicitud, hoy),
    FECHA_ENTREGA:          safeStr(s.fechaEntrega, hoy),
    // Tabla de materiales — columnas oficiales: RESERVA | STOCK | CANTIDAD | DESCRIPCIÓN
    MATERIALES: (s.items || []).map(i => ({
      RESERVA:     safeStr(i.sapCode, ''),   // Código SAP
      STOCK:       safeStr(i.axCode, ''),    // Código AX
      CANTIDAD:    safeNum(i.cantidad),
      DESCRIPCION: safeStr(i.nombre || i.name, ''),
      // Seriales — sección aparte en el documento oficial
      _requiereSerial: !!i.requiereSerial,
      _modoSerial:     i.modoSerial || 'individual',
      _seriales:       i.seriales || [],
      _serialInicio:   i.serialInicio || '',
      _serialFin:      i.serialFin || '',
    })),
    // Seriales (solo materiales que lo requieren)
    SERIALES: (s.items || []).filter(i => i.requiereSerial),
  };
}

function showMemo(s) {
  // ── Transformar al formato oficial ──
  const memo = mapearParaMemo(s);

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Memo de despacho</h2>
      <button id="fm-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div id="memo-body" class="px-5 py-4 space-y-4 text-sm">

      <!-- Encabezado corporativo -->
      <div class="text-center border-b border-gray-200 pb-3">
        <p class="font-bold text-sm uppercase">DISTRIBUIDORA DE ELECTRICIDAD DELSUR S.A. DE C.V.</p>
        <p class="text-xs text-gray-500 mt-0.5 uppercase">OTC - GERENCIA COMERCIAL</p>
        <p class="text-xs font-semibold text-gray-700 mt-1 uppercase tracking-wide">DESPACHO / CARGA DE MATERIALES</p>
      </div>

      <!-- Campos del encabezado — orden y nombres exactos del documento oficial -->
      <table class="w-full text-xs border-collapse">
        <tr>
          <td class="py-1 pr-2 text-gray-500 uppercase font-semibold w-40">USUARIO RESPONSABLE:</td>
          <td class="py-1 border-b border-gray-400 font-medium text-gray-900">${memo.USUARIO_RESPONSABLE}</td>
        </tr>
        <tr><td colspan="2" class="py-0.5"></td></tr>
        <tr>
          <td class="py-1 pr-2 text-gray-500 uppercase font-semibold">EMPRESA CONTRATISTA:</td>
          <td class="py-1 border-b border-gray-400 font-medium text-gray-900">${memo.EMPRESA_CONTRATISTA}</td>
        </tr>
        <tr><td colspan="2" class="py-0.5"></td></tr>
        <tr>
          <td class="py-1 pr-2 text-gray-500 uppercase font-semibold">INSTALADOR RESPONSABLE:</td>
          <td class="py-1 border-b border-gray-400 font-medium text-gray-900">${memo.INSTALADOR_RESPONSABLE}</td>
        </tr>
        <tr><td colspan="2" class="py-0.5"></td></tr>
        <tr>
          <td class="py-1 pr-2 text-gray-500 uppercase font-semibold">ENTREGADO POR:</td>
          <td class="py-1 border-b border-gray-400 font-medium text-gray-900">${memo.ENTREGADO_POR}</td>
        </tr>
        <tr><td colspan="2" class="py-0.5"></td></tr>
        <tr>
          <td class="py-1 pr-2 text-gray-500 uppercase font-semibold">PLACA DE VEHICULO:</td>
          <td class="py-1 border-b border-gray-400 font-mono font-medium text-gray-900">${memo.PLACA_VEHICULO}</td>
        </tr>
        <tr><td colspan="2" class="py-0.5"></td></tr>
        <tr>
          <td class="py-1 pr-2 text-gray-500 uppercase font-semibold">FECHA DE SOLICITUD:</td>
          <td class="py-1 border-b border-gray-400 font-medium text-gray-900">${memo.FECHA_SOLICITUD}</td>
        </tr>
        <tr><td colspan="2" class="py-0.5"></td></tr>
        <tr>
          <td class="py-1 pr-2 text-gray-500 uppercase font-semibold">FECHA ENTREGA DE MATERIAL:</td>
          <td class="py-1 border-b border-gray-400 font-medium text-gray-900">${memo.FECHA_ENTREGA}</td>
        </tr>
      </table>

      <!-- Tabla de materiales — columnas del documento oficial -->
      <div>
        <table class="w-full text-xs border border-gray-300">
          <thead>
            <tr class="bg-gray-100">
              <th class="border border-gray-300 px-2 py-1.5 text-center uppercase font-semibold text-gray-700">RESERVA</th>
              <th class="border border-gray-300 px-2 py-1.5 text-center uppercase font-semibold text-gray-700">STOCK</th>
              <th class="border border-gray-300 px-2 py-1.5 text-center uppercase font-semibold text-gray-700">CANTIDAD</th>
              <th class="border border-gray-300 px-2 py-1.5 text-left uppercase font-semibold text-gray-700">DESCRIPCIÓN</th>
            </tr>
          </thead>
          <tbody>
            ${memo.MATERIALES.map(i => `
              <tr>
                <td class="border border-gray-300 px-2 py-1.5 text-center font-mono text-gray-800">${i.RESERVA}</td>
                <td class="border border-gray-300 px-2 py-1.5 text-center font-mono text-gray-800">${i.STOCK}</td>
                <td class="border border-gray-300 px-2 py-1.5 text-center font-semibold text-gray-900">${i.CANTIDAD}</td>
                <td class="border border-gray-300 px-2 py-1.5 text-gray-900 uppercase">${i.DESCRIPCION}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>

      <!-- Seriales de medidores / sellos — sección oficial -->
      ${memo.SERIALES.length > 0 ? `
      <div>
        <p class="text-xs font-semibold uppercase text-gray-600 mb-2">Serial de medidores / sellos entregados</p>
        <table class="w-full text-xs border border-gray-300">
          <thead>
            <tr class="bg-gray-100">
              <th class="border border-gray-300 px-2 py-1.5 text-left uppercase font-semibold text-gray-700">STOCK</th>
              <th class="border border-gray-300 px-2 py-1.5 text-left uppercase font-semibold text-gray-700">RESERVA</th>
              <th class="border border-gray-300 px-2 py-1.5 text-left uppercase font-semibold text-gray-700">DESCRIPCIÓN</th>
              <th class="border border-gray-300 px-2 py-1.5 text-center uppercase font-semibold text-gray-700">CANTIDAD</th>
              <th class="border border-gray-300 px-2 py-1.5 text-center uppercase font-semibold text-gray-700">INICIO</th>
              <th class="border border-gray-300 px-2 py-1.5 text-center uppercase font-semibold text-gray-700">FIN</th>
            </tr>
          </thead>
          <tbody>
            ${memo.SERIALES.map(i => {
              const axCode  = safeStr(i.axCode,'');
              const sapCode = safeStr(i.sapCode,'');
              const desc    = safeStr(i.nombre||i.name,'').toUpperCase();
              const cant    = safeNum(i.cantidad);
              const inicio  = i.modoSerial === 'rango' ? safeStr(i.serialInicio,'') : (i.seriales||[])[0] || '';
              const fin     = i.modoSerial === 'rango' ? safeStr(i.serialFin,'')   : (i.seriales||[]).slice(-1)[0] || '';
              return `<tr>
                <td class="border border-gray-300 px-2 py-1.5 font-mono text-gray-800">${axCode}</td>
                <td class="border border-gray-300 px-2 py-1.5 font-mono text-gray-800">${sapCode}</td>
                <td class="border border-gray-300 px-2 py-1.5 text-gray-900 uppercase">${desc}</td>
                <td class="border border-gray-300 px-2 py-1.5 text-center font-semibold">${cant}</td>
                <td class="border border-gray-300 px-2 py-1.5 text-center font-mono">${inicio}</td>
                <td class="border border-gray-300 px-2 py-1.5 text-center font-mono">${fin}</td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
        ${memo.SERIALES.some(i => i.modoSerial === 'individual' && (i.seriales||[]).length > 1) ? `
        <div class="mt-2 space-y-2">
          ${memo.SERIALES.filter(i => i.modoSerial === 'individual' && (i.seriales||[]).length > 1).map(i => `
            <div class="border border-gray-200 rounded p-2">
              <p class="text-xs font-semibold text-gray-600 uppercase mb-1">${safeStr(i.nombre||i.name).toUpperCase()}</p>
              <div class="flex flex-wrap gap-1">
                ${(i.seriales||[]).map(ser => `<span class="text-xs px-1.5 py-0.5 border border-gray-300 rounded font-mono text-gray-700">${ser}</span>`).join('')}
              </div>
            </div>`).join('')}
        </div>` : ''}
      </div>` : ''}

      <!-- Espacios de firma — exactamente como en el documento oficial -->
      <div class="grid grid-cols-2 gap-6 pt-4">
        <div class="text-center space-y-1">
          <div class="border-b border-gray-500 h-10"></div>
          <p class="text-xs uppercase font-semibold text-gray-500">FIRMA DE ENTREGADO</p>
        </div>
        <div class="text-center space-y-1">
          <div class="border-b border-gray-500 h-10"></div>
          <p class="text-xs uppercase font-semibold text-gray-500">FIRMA DE RECIBIDO</p>
        </div>
      </div>

    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fm-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cerrar</button>
      <button id="fm-print"  class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">🖨️ Imprimir memo</button>
    </div>`);

  ov.querySelector('#fm-close').onclick = ov.querySelector('#fm-cancel').onclick = () => ov.remove();

  // Botón descargar JSON (para usar con el script Python en PC)
  ov.querySelector('#fm-export-json')?.addEventListener('click', () => {
    const datos = {
      USUARIO_RESPONSABLE:    memo.USUARIO_RESPONSABLE,
      EMPRESA_CONTRATISTA:    memo.EMPRESA_CONTRATISTA,
      INSTALADOR_RESPONSABLE: memo.INSTALADOR_RESPONSABLE,
      ENTREGADO_POR:          memo.ENTREGADO_POR,
      PLACA_VEHICULO:         memo.PLACA_VEHICULO,
      FECHA_SOLICITUD:        memo.FECHA_SOLICITUD,
      FECHA_ENTREGA:          memo.FECHA_ENTREGA,
      CANTIDADES: Object.fromEntries(
        memo.MATERIALES.map(i => [i.RESERVA, String(i.CANTIDAD)])
      ),
    };
    const blob = new Blob([JSON.stringify(datos, null, 2)], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `despacho-${memo.FECHA_ENTREGA || 'datos'}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 3000);
  });

  ov.querySelector('#fm-print').onclick = () => imprimirDespacho(memo);
}

// ─────────────────────────────────────────
// DESCARGAR JSON PARA SCRIPT PYTHON
// ─────────────────────────────────────────
function descargarMemoJSON(memo) {
  // Construir objeto JSON compatible con generar_memo.py
  const datos = {
    usuarioResponsable:    memo.USUARIO_RESPONSABLE    || '',
    empresaContratista:    memo.EMPRESA_CONTRATISTA    || '',
    instaladorResponsable: memo.INSTALADOR_RESPONSABLE || '',
    entregadoPor:          memo.ENTREGADO_POR          || '',
    placaVehiculo:         memo.PLACA_VEHICULO         || '',
    fechaSolicitud:        memo.FECHA_SOLICITUD        || '',
    fechaEntrega:          memo.FECHA_ENTREGA          || '',
    items: (memo.MATERIALES || []).map(function(i) {
      return {
        sapCode:  i.RESERVA  || '',
        axCode:   i.STOCK    || '',
        nombre:   i.NOMBRE   || '',
        cantidad: i.CANTIDAD || '',
        unit:     i.UNIDAD   || '',
      };
    }),
    seriales: (memo.SERIALES || []).map(function(i) {
      return {
        sapCode:      i.sapCode   || i.RESERVA || '',
        axCode:       i.axCode    || i.STOCK   || '',
        nombre:       i.nombre    || i.NOMBRE  || '',
        cantidad:     i.cantidad  || '',
        modoSerial:   i.modoSerial || 'individual',
        seriales:     i.seriales   || [],
        serialInicio: i.serialInicio || '',
        serialFin:    i.serialFin    || '',
      };
    }),
  };

  const json    = JSON.stringify(datos, null, 2);
  const blob    = new Blob([json], { type: 'application/json' });
  const url     = URL.createObjectURL(blob);
  const a       = document.createElement('a');
  const fecha   = (datos.fechaEntrega || new Date().toISOString().slice(0,10)).replace(/-/g,'');
  const usuario = (datos.usuarioResponsable || 'despacho').replace(/\s/g,'_');
  a.href        = url;
  a.download    = 'despacho_' + usuario + '_' + fecha + '.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
  showToast('JSON descargado. Cópialo al escritorio y corre el script.', 'success');
}

// ─────────────────────────────────────────
// IMPRIMIR DESPACHO OFICIAL — 2 páginas
// ─────────────────────────────────────────

// 6 bloques fijos de la página 2
const BLOQUES_SERIALES = [
  { ax:'700101', sap:'200129', nombre:'MEDIDOR BIFILAR DOMICILIAR BASE A, 100 A (ETE1-330)',                                          filas:30, tipo:'serial' },
  { ax:'700102', sap:'355518', nombre:'MEDIDOR TRIFILAR DOMICILIAR BASE A, 100 A (ETE1-330)',                                         filas:30, tipo:'serial' },
  { ax:'400931', sap:'354549', nombre:'SELLO ACRILICO VERDE (SERV. NVOS., MTTO.) (CABLE 30 CM) (FTMED-30)',                           filas:10, tipo:'sello'  },
  { ax:'700326', sap:'338362', nombre:'MEDIDOR FORMA 2(S) T/ESPIGA, CLASE 100, TRIFI. 240 V, 15/100',                                 filas:5,  tipo:'serial' },
  { ax:'700332', sap:'355064', nombre:'MEDIDOR FORMA 16s, CLASE 200, 120-277V. 8 CANALES DE MEM. 200 AMP. C/BASE 7 TERMI. (ETE-16s)', filas:5,  tipo:'serial' },
  { ax:'700333', sap:'338357', nombre:'MEDIDOR FORMA 12s CLASE 200, 120-277V, TRIFILAR, 60Hz, 8 Canales de memoria, C/Base 5 Term (ETE-12s)', filas:3, tipo:'serial' },
];

const FILAS_DOC = [
  {sap:'RESERVA',ax:'STOCK',desc:'DESCRIPICIÓN',header:'col'},
  {sap:'USO HABITUAL',ax:'',desc:'',header:'sec'},
  {sap:'221477',ax:'50203',desc:'ALAMBRE COBRE THHN 8 AWG 600 V FORRO PLASTICO'},
  {sap:'213719',ax:'50806',desc:'CABLE DUPLEX AL #6 ACSR SETTER'},
  {sap:'328541',ax:'50807',desc:'CABLE TRIPLEX AL. #6 ACSR PALUDINA'},
  {sap:'352453',ax:'250201',desc:'CONECTOR DE COMPRESIÓN YPC2A8U'},
  {sap:'352460',ax:'250202',desc:'CONECTOR DE COMPRESIÓN YPC26R8U'},
  {sap:'352461',ax:'250203',desc:'CONECTOR DE COMPRESIÓN YP2U3'},
  {sap:'352462',ax:'250204',desc:'CONECTOR DE COMPRESIÓN YP26AU2'},
  {sap:'353112',ax:'400910',desc:'ANCLA PLASTICA 1 1/2 X 7 (FTN1-120)'},
  {sap:'354045',ax:'400919',desc:'TORNILLO CABEZA PLANA DE 11/2 PLG X 7MM (Gruesa de 144 unidades)'},
  {sap:'354549',ax:'400931',desc:'SELLO ACRILICO VERDE (SERV. NVOS., MTTO.) (CABLE 30 CM) (FTMED-30)'},
  {sap:'200129',ax:'700101',desc:'MEDIDOR BIFILAR DOMICILIAR BASE A, 100 A (ETE1-330)'},
  {sap:'355518',ax:'700102',desc:'MEDIDOR TRIFILAR DOMICILIAR BASE A, 100 A (ETE1-330)'},
  {sap:'338362',ax:'700326',desc:'MEDIDOR FORMA 2(S) T/ESPIGA, CLASE 100, TRIFI. 240 V, 15/100'},
  {sap:'219359',ax:'750109',desc:'CINTA AISLANTE SUPER 3M #33'},
  {sap:'MATERIAL PARA CL200',ax:'',desc:'',header:'sec'},
  {sap:'328560',ax:'50205',desc:'CABLE COBRE THHN # 2 AWG 600 V FORRO PLASTICO (ETM3-310)'},
  {sap:'243940',ax:'50209',desc:'CABLE COBRE THHN # 1/0 AWG 19 HILOS (ETM3-310)'},
  {sap:'337775',ax:'250101',desc:'CONECTOR MECANICO PERNO PARTIDO KSU-23'},
  {sap:'337776',ax:'250102',desc:'CONECTOR MECANICO PERNO PARTIDO KSU-26'},
  {sap:'337777',ax:'250103',desc:'CONECTOR MECANICO PERNO PARTIDO KSU-29'},
  {sap:'210525',ax:'400833',desc:'TUBO EMT 2 PLG ALUMINIO UL (FTN1-700)'},
  {sap:'212720',ax:'400834',desc:'CUERPO TERMINAL PARA EMT DE 2 PLG CON ABRAZADERA'},
  {sap:'214221',ax:'400836',desc:'CONECTOR EMT DE 2 PLG CON TORNILLO'},
  {sap:'301560',ax:'400837',desc:'GRAPA CONDUIT EMT DE 2 PLG'},
  {sap:'245979',ax:'400838',desc:'TUBO EMT DE 2 1/2 PLG (FTN1-700)'},
  {sap:'355070',ax:'400839',desc:'CUERPO TERMINAL PARA EMT DE 2 1/2 PLG CON ABRAZADERA'},
  {sap:'244569',ax:'400840',desc:'CONECTOR EMT DE 2 1/2 PLG CON TORNILLO'},
  {sap:'353992',ax:'400841',desc:'GRAPA CONDUIT EMT DE 2 1/2 PLG'},
  {sap:'211373',ax:'400903',desc:'CINTA BAND IT 3/4'},
  {sap:'211375',ax:'400904',desc:'HEBILLA PARA CINTA BAND IT 3/4'},
  {sap:'353121',ax:'700406',desc:'TAPADERA DE POLICARBONATO PARA BASES SOCKET TIPO ESPIGA'},
  {sap:'355064',ax:'700332',desc:'MEDIDOR FORMA 16s, CLASE 200, 120-277V. 8 CANALES DE MEM. 200 AMP. C/BASE 7 TERMI. (ETE-16s)'},
  {sap:'338357',ax:'700333',desc:'MEDIDOR FORMA 12s CLASE 200, 120-277V, TRIFILAR, 60Hz, 8 Canales de memoria, C/Base 5 Term (ETE-12s)'},
  {sap:'353099',ax:'401102',desc:'BASES SOCKETS DE 5 TERMINALES, 200 AMP'},
  {sap:'353110',ax:'700404',desc:'BASES SOCKETS 100 AMP. P/MED ESPIGA F 2(S)'},
  {sap:'353088',ax:'401101',desc:'BASES SOCKETS DE 7 TERMINALES MARCA THOMAS & BETTS, 200 AMPS'},
  {sap:'ACOMETIDA ESPECIAL',ax:'',desc:'',header:'sec'},
  {sap:'200468',ax:'50401',desc:'CABLE DE ALUMINIO #2 AWG F.P. 7 HILOS (ETM3-330)'},
  {sap:'200472',ax:'50809',desc:'CABLE DE ALUMINIO ACSR DESNUDO #2 AWG 7 HILOS SPARROW (ETM3-350)'},
  {sap:'200469',ax:'50402',desc:'CABLE DE ALUMINIO #1/0 AWG F.P. 7 HILOS (ETM3-330)'},
  {sap:'200473',ax:'50810',desc:'CABLE DE ALUMINIO ACSR DESNUDO #1/0 AWG 7 HILOS, (ETM3-350)'},
  {sap:'213410',ax:'50505',desc:'CABLE TRIPLEX ACSR 1/0 AWG NERITINA (ETM3-330)'},
  {sap:'214726',ax:'150202',desc:'REMATE PREFORMADO ACSR 2 AWG (ETM1-240)'},
  {sap:'214727',ax:'150205',desc:'REMATE PREFORMADO ACSR 1/0 AWG (ETM1-240)'},
  {sap:'352463',ax:'250205',desc:'CONECTOR MECANICO COMPRESIÓN YP25U25'},
  {sap:'SUBTERRÁNEO',ax:'',desc:'',header:'sec'},
  {sap:'219527',ax:'250418',desc:'TERMINAL DE OJO CABLE 4 AWG, DIAMETRO 3/8 PLG (ETM1-460)'},
  {sap:'221062',ax:'250420',desc:'TERMINAL DE OJO 1/0 DIAMETRO 3/8 PLG (FTN1-320)'},
  {sap:'200367',ax:'50220',desc:'CABLE COBRE XHHW # 4 AWG 19 HILOS (ETM3-310)'},
  {sap:'282485',ax:'50219',desc:'CABLE COBRE XHHW # 6 AWG 7 HILOS (ETM3-310)'},
  {sap:'350560',ax:'50231',desc:'CABLE COBRE RHHW #1/0 AWG'},
  {sap:'212896',ax:'250419',desc:'TERMINAL DE OJO NO. 2 DIAMETRO 3/8 PLG'},
  {sap:'350564',ax:'50234',desc:'CABLE COBRE RHHW # 2 AWG 19 HILOS'},
  {sap:'PATRON ANTIHURTO Y TELEGESTIÓN',ax:'',desc:'',header:'sec'},
  {sap:'221472',ax:'50710',desc:'CABLE CONCENTRICO TELESCOPICO CCA BIFILAR 6 AWG PARA ACOMETIDA AEREA (ETM3-470)'},
  {sap:'200413',ax:'50721',desc:'CABLE CONCENTRICO TELESCOPICO CCA TRIFILAR 6 AWG PARA ACOMETIDA AEREA (ETM3-470)'},
  {sap:'211829',ax:'400707',desc:'CAJA TRANSPARENTE DE POLICARBONATO PARA MEDIDOR TOTALIZADOR'},
  {sap:'213340',ax:'150419',desc:'PINZAS DE RETENCION PARA CABLE CONCENTRICO TELESCOPICO #1/0 AWG (FTN1-760)'},
  {sap:'222315',ax:'150420',desc:'PINZAS DE RETENCION PARA CABLE CONCENTRICO TELESCOPICO DE ACOMETIDA BT #6 (FTN1-770)'},
  {sap:'353730',ax:'750116',desc:'CINTA DE BLINDAJE 3M, CAT. ARMORCAST 4560-10'},
  {sap:'338363',ax:'700340',desc:'MEDIDOR TELEGESTIONADO BASE A - 240V (FTMED-32)'},
  {sap:'338361',ax:'700339',desc:'MEDIDOR RESIDENCIAL PREPAGO CLASE 100, 120V CON TELEGESTION'},
  {sap:'338360',ax:'700338',desc:'MEDIDOR RESIDENCIAL POSTPAGO CLASE 100, 120-240V CON TELEGESTION'},
];

function imprimirDespacho(memo) {
  const cantMap = {};
  for (const it of (memo.MATERIALES || [])) {
    cantMap[String(it.RESERVA).trim()] = it.CANTIDAD;
  }

  // Mapa de seriales por SAP — rangos se expanden a lista individual
  const serialMap = {};
  for (const it of (memo.MATERIALES || [])) {
    const sap = String(it.RESERVA || '').trim();
    if (it._seriales && it._seriales.length > 0) {
      serialMap[sap] = { tipo:'individual', seriales: it._seriales };
    } else if (it._serialInicio) {
      const ini    = String(it._serialInicio).trim();
      const fin    = String(it._serialFin || '').trim();
      const nIni   = parseInt(ini.replace(/[^0-9]/g,''), 10);
      const nFin   = parseInt(fin.replace(/[^0-9]/g,''), 10);
      const prefix = ini.replace(/[0-9]+$/, '');
      if (!isNaN(nIni) && !isNaN(nFin) && nFin >= nIni && (nFin - nIni) <= 500) {
        const expanded = [];
        const digits   = String(nFin).length;
        for (let n = nIni; n <= nFin; n++) {
          expanded.push(prefix + String(n).padStart(digits, '0'));
        }
        serialMap[sap] = { tipo:'individual', seriales: expanded };
      } else {
        serialMap[sap] = { tipo:'rango', inicio: ini, fin: fin };
      }
    }
  }

  // ── CSS compartido ──
  const css = `
    @page { size:215.9mm 279.4mm; margin:8.8mm 6.3mm 4.9mm 12.7mm; }
    *{margin:0;padding:0;box-sizing:border-box;}
    body{font-family:Arial,sans-serif;font-size:6pt;color:#000;background:#fff;width:196.9mm;}
    .empresa{font-size:7pt;font-weight:bold;}
    .sub{font-size:6pt;}
    .lbl{font-size:5.5pt;display:block;}
    .linea{font-size:6pt;font-weight:bold;display:inline-block;
      border-bottom:0.4pt solid #000;min-height:2.5mm;padding-bottom:0.2mm;
      min-width:45mm;max-width:90mm;}
    .tm{width:196.9mm;border-collapse:collapse;font-size:5.5pt;margin-top:1.5mm;table-layout:fixed;}
    .tm td,.tm th{border:0.4pt solid #000;padding:0.25mm 0.4mm;vertical-align:middle;overflow:hidden;line-height:1.35;}
    .c-sap{width:17mm;}.c-ax{width:11mm;}.c-desc{width:155mm;}.c-cant{width:13.9mm;}
    .th{font-weight:bold;font-size:6pt;text-align:center;}
    .sec{font-weight:bold;text-align:center;}
    .code{text-align:center;font-size:5pt;}
    .cant{text-align:center;font-weight:bold;}
    /* Página 2 */
    .page2{page-break-before:always;}
    .titulo-p2{font-size:7pt;font-weight:bold;margin-bottom:2mm;}
    .pg2{position:relative;width:196.9mm;height:260mm;}
    .tb{position:absolute;border:0.5pt solid #000;overflow:hidden;display:flex;flex-direction:column;}
    .tb-hdr{background:#F7C6AC;font-weight:bold;border-bottom:0.5pt solid #000;flex-shrink:0;}
    .tb-hdr table{width:100%;border-collapse:collapse;table-layout:fixed;}
    .tb-hdr .cod{font-size:5pt;padding:0.5mm 0.8mm;border-right:0.5pt solid #000;vertical-align:middle;text-align:center;width:15mm;min-width:15mm;max-width:15mm;}
    .tb-hdr .nom{font-size:5.5pt;padding:0.5mm 0.8mm;vertical-align:middle;text-align:center;}
    .tb-body{flex:1;min-height:0;overflow:hidden;}
    .tb-body table{width:100%;height:100%;border-collapse:collapse;table-layout:fixed;}
    .tb-body td{border:0.3pt solid #ccc;padding:0 0.5mm;font-size:5pt;}
    .nb{width:15mm;min-width:15mm;max-width:15mm;text-align:center;color:#555;border-right:0.4pt solid #999;}
    .hc{background:#f5f5f5;font-weight:bold;text-align:center;font-size:4.5pt;border-bottom:0.4pt solid #999;}
    .cc{text-align:center;border-right:0.3pt solid #999;}
    .ci{text-align:center;border-right:0.3pt solid #999;}
    .cf{text-align:center;}
    .filled{color:#000;font-weight:bold;}
    @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
  `;

  // ── PÁGINA 1: materiales ──
  const filas = FILAS_DOC.map(function(row) {
    if (row.header === 'col') return '<tr><th class="th">RESERVA</th><th class="th">STOCK</th><th class="th">DESCRIPICIÓN</th><th class="th cant">CANTIDAD</th></tr>';
    if (row.header === 'sec') return '<tr><td class="sec" colspan="4">' + row.sap + '</td></tr>';
    const cant = cantMap[row.sap] || '';
    return '<tr><td class="code">' + row.sap + '</td><td class="code">' + row.ax + '</td><td>' + row.desc + '</td><td class="cant">' + cant + '</td></tr>';
  }).join('');

  const v = memo;
  const p1 =
    '<table style="width:196.9mm;border-collapse:collapse;margin-bottom:1.5mm;">' +
      '<colgroup><col style="width:98mm"><col style="width:98.9mm"></colgroup>' +
      '<tr>' +
        '<td rowspan="2" style="vertical-align:top;padding-right:2mm;border:none;">' +
          '<div class="empresa">DISTRUIBUIDORA DE ELECTRICIDAD DELSUR S.A. DE C.V.</div>' +
          '<div class="sub">OTC - GERENCIA COMERCIAL</div>' +
          '<div class="sub">DESPACHO/ CARGA DE MATERIALES</div>' +
        '</td>' +
        '<td style="border:none;padding-bottom:1.5mm;"><span class="lbl">USUARIO RESPONSABLE:</span><span class="linea">' + (v.USUARIO_RESPONSABLE||'') + '</span></td>' +
      '</tr>' +
      '<tr><td style="border:none;padding-bottom:1.5mm;"><span class="lbl">INSTALADOR RESPONSABLE:</span><span class="linea">' + (v.INSTALADOR_RESPONSABLE||'') + '</span></td></tr>' +
      '<tr>' +
        '<td style="border:none;padding-top:2mm;padding-bottom:1.5mm;"><span class="lbl">EMPRESA CONTRATISTA:</span><span class="linea">' + (v.EMPRESA_CONTRATISTA||'') + '</span></td>' +
        '<td style="border:none;padding-top:2mm;padding-bottom:1.5mm;"><span class="lbl">FIRMA DE RECIBIDO:</span><span class="linea">&nbsp;</span></td>' +
      '</tr>' +
      '<tr>' +
        '<td style="border:none;padding-bottom:1.5mm;"><span class="lbl">ENTREGADO POR:</span><span class="linea">' + (v.ENTREGADO_POR||'') + '</span></td>' +
        '<td style="border:none;padding-bottom:1.5mm;"><span class="lbl">PLACA DE VEHICULO:</span><span class="linea">' + (v.PLACA_VEHICULO||'') + '</span></td>' +
      '</tr>' +
      '<tr>' +
        '<td style="border:none;padding-bottom:1.5mm;"><span class="lbl">FIRMA DE ENTREGADO:</span><span class="linea">&nbsp;</span></td>' +
        '<td style="border:none;padding-bottom:1.5mm;"><span class="lbl">FECHA ENTREGA DE MATERIAL:</span><span class="linea">' + (v.FECHA_ENTREGA||'') + '</span></td>' +
      '</tr>' +
      '<tr>' +
        '<td style="border:none;"><span class="lbl">FECHA DE SOLICITUD</span><span class="linea">' + (v.FECHA_SOLICITUD||'') + '</span></td>' +
        '<td style="border:none;"></td>' +
      '</tr>' +
    '</table>' +
    '<table class="tm"><colgroup><col class="c-sap"><col class="c-ax"><col class="c-desc"><col class="c-cant"></colgroup>' + filas + '</table>';

  // ── PÁGINA 2: seriales ──
  function buildBloque(b) {
    const serData = serialMap[b.sap] || null;
    const isSello = b.tipo === 'sello';
    const ancho   = b.filas >= 30 ? '62mm' : b.filas >= 10 ? '62mm' : '62mm';

    let filasHtml = '';
    if (isSello) {
      // Encabezado de columnas
      filasHtml += '<tr class="sub-header"><td class="num"></td><td class="col-cant">Cantidad</td><td class="col-ini">Inicio</td><td class="col-fin">Fin</td></tr>';
      for (let i = 1; i <= b.filas; i++) {
        let cant = '', ini = '', fin = '';
        if (serData && serData.tipo === 'rango' && i === 1) {
          cant = cantMap[b.sap] || '';
          ini  = serData.inicio || '';
          fin  = serData.fin    || '';
        }
        const cls = cant ? ' class="filled"' : '';
        filasHtml += '<tr><td class="num">' + i + '</td>' +
          '<td class="col-cant"' + cls + '>' + cant + '</td>' +
          '<td class="col-ini"'  + cls + '>' + ini  + '</td>' +
          '<td class="col-fin"'  + cls + '>' + fin  + '</td></tr>';
      }
    } else {
      // Filas con serial individual
      const seriales = serData && serData.tipo === 'individual' ? serData.seriales : [];
      for (let i = 1; i <= b.filas; i++) {
        const val = seriales[i-1] || '';
        const cls = val ? ' class="filled"' : '';
        filasHtml += '<tr><td class="num">' + i + '</td><td' + cls + '>' + val + '</td></tr>';
      }
    }

    return '<div class="bloque" style="width:' + ancho + '">' +
      '<div class="bloque-header">' +
        '<div class="codigos">' + b.ax + '<br>' + b.sap + '</div>' +
        '<div class="nombre">' + b.nombre + '</div>' +
      '</div>' +
      '<div class="bloque-filas"><table>' + filasHtml + '</table></div>' +
    '</div>';
  }

  function buildHdrTd(b) {
    return '<div class="tb-hdr"><table><tr>' +
      '<td class="cod">' + b.ax + '<br>' + b.sap + '</td>' +
      '<td class="nom">' + b.nombre + '</td>' +
    '</tr></table></div>';
  }

  function buildFilas(b) {
    const serData = serialMap[b.sap] || null;
    let rows = '';
    if (b.tipo === 'sello') {
      rows += '<tr><td class="nb hc"></td><td class="cc hc">Cantidad</td><td class="ci hc">Inicio</td><td class="cf hc">Fin</td></tr>';
      for (let i = 1; i <= b.filas; i++) {
        let cant='', ini='', fin='';
        if (serData && serData.tipo==='rango' && i===1) { cant=cantMap[b.sap]||''; ini=serData.inicio||''; fin=serData.fin||''; }
        const cls = cant ? ' class="filled"' : '';
        rows += '<tr><td class="nb">' + i + '</td><td class="cc"' + cls + '>' + cant + '</td><td class="ci"' + cls + '>' + ini + '</td><td class="cf"' + cls + '>' + fin + '</td></tr>';
      }
    } else {
      const sers = serData && serData.tipo==='individual' ? serData.seriales : [];
      for (let i = 1; i <= b.filas; i++) {
        const val = sers[i-1] || '';
        const cls = val ? ' class="filled"' : '';
        rows += '<tr><td class="nb">' + i + '</td><td' + cls + '>' + val + '</td></tr>';
      }
    }
    // table height:100% + flex:1 on tb-body fills rows automatically
    return '<div class="tb-body"><table>' + rows + '</table></div>';
  }

  const p2 =
    '<div class="page2">' +
      '<div class="titulo-p2">Serial de medidores /sellos entregados</div>' +
      '<div class="pg2">' +
        // TB0: Bifilar
        '<div class="tb" style="left:0;top:0;width:62.6mm;height:166.4mm;">' + buildHdrTd(BLOQUES_SERIALES[0]) + buildFilas(BLOQUES_SERIALES[0]) + '</div>' +
        // TB1: Trifilar
        '<div class="tb" style="left:69.2mm;top:0;width:62.9mm;height:166.4mm;">' + buildHdrTd(BLOQUES_SERIALES[1]) + buildFilas(BLOQUES_SERIALES[1]) + '</div>' +
        // TB2: Sello
        '<div class="tb tb-sello" style="left:138.1mm;top:0;width:50.4mm;height:59.7mm;">' + buildHdrTd(BLOQUES_SERIALES[2]) + buildFilas(BLOQUES_SERIALES[2]) + '</div>' +
        // TB3: Forma 2S
        '<div class="tb tb-small" style="left:0;top:172.3mm;width:62.6mm;height:36.1mm;">' + buildHdrTd(BLOQUES_SERIALES[3]) + buildFilas(BLOQUES_SERIALES[3]) + '</div>' +
        // TB4: Forma 16s
        '<div class="tb tb-small" style="left:69.2mm;top:172.3mm;width:62.9mm;height:36.1mm;">' + buildHdrTd(BLOQUES_SERIALES[4]) + buildFilas(BLOQUES_SERIALES[4]) + '</div>' +
        // Tabla Forma 12s
        '<div class="tb tb-tiny" style="left:0;top:215mm;width:82.3mm;height:36.1mm;">' + buildHdrTd(BLOQUES_SERIALES[5]) + buildFilas(BLOQUES_SERIALES[5]) + '</div>' +
      '</div>' +
    '</div>';

  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Memo Despacho</title><style>' + css + '</style></head><body>' + p1 + p2 + '</body></html>';

  // Imprimir via iframe
  let ifr = document.getElementById('__print_frame');
  if (ifr) ifr.remove();
  ifr = document.createElement('iframe');
  ifr.id = '__print_frame';
  ifr.style.cssText = 'position:fixed;top:-9999px;left:-9999px;width:216mm;height:280mm;border:none;';
  document.body.appendChild(ifr);
  const iDoc = ifr.contentDocument || ifr.contentWindow.document;
  iDoc.open(); iDoc.write(html); iDoc.close();
  setTimeout(function() {
    try { ifr.contentWindow.print(); }
    catch(e) {
      const w = window.open('','_blank');
      if (w) { w.document.write(html); w.document.close(); }
    }
  }, 500);
}

// ─────────────────────────────────────────
// IMPORTAR DESDE EXCEL
// ─────────────────────────────────────────
async function showImportExcel(db, session, existingItems, refreshFn) {
  if (!window.XLSX) {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js';
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  let toImport = [], toSkip = [];

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Importar desde Excel</h2>
      <button id="ix-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div id="ix-step-file" class="px-5 py-4 space-y-3">
      <p class="text-sm text-gray-500">Columnas requeridas: <strong>SAP</strong>, <strong>AX</strong>, <strong>Nombre material</strong></p>
      <label class="flex flex-col items-center justify-center border-2 border-dashed border-gray-300 rounded-xl p-5 cursor-pointer hover:border-blue-400 transition-colors">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#9CA3AF" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" class="mb-2">
          <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
          <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
        <p class="text-sm text-gray-600 font-medium">Seleccionar archivo Excel</p>
        <input id="ix-file" type="file" accept=".xlsx,.xls" class="hidden"/>
      </label>
      <div id="ix-file-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2"></div>
    </div>
    <div id="ix-step-preview" class="hidden px-5 py-4 space-y-3">
      <div id="ix-summary" class="grid grid-cols-3 gap-2"></div>
      <div id="ix-tabla" class="rounded-lg border border-gray-200 overflow-hidden max-h-52 overflow-y-auto text-xs"></div>
      <div id="ix-skipped" class="hidden"></div>
      <div id="ix-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2"></div>
      <div class="flex gap-3 pt-1">
        <button id="ix-back" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm">← Volver</button>
        <button id="ix-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background:#1B4F8A">Importar</button>
      </div>
    </div>
    <div id="ix-step-result" class="hidden px-5 py-6 text-center space-y-3">
      <div class="w-14 h-14 rounded-full flex items-center justify-center mx-auto" style="background:#DCFCE7">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#166534" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
      </div>
      <p class="font-semibold text-gray-900">Importación completada</p>
      <div id="ix-result-txt" class="space-y-1"></div>
      <button id="ix-done" class="text-white font-medium rounded-lg px-6 py-2.5 text-sm" style="background:#1B4F8A">Cerrar</button>
    </div>`);

  ov.querySelector('#ix-close').onclick = () => ov.remove();

  ov.querySelector('#ix-file').addEventListener('change', async (e) => {
    const file  = e.target.files[0];
    const errEl = ov.querySelector('#ix-file-err');
    errEl.classList.add('hidden');
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const wb   = window.XLSX.read(data, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = window.XLSX.utils.sheet_to_json(ws, { defval: null });
      if (rows.length === 0) throw new Error('El archivo está vacío.');

      const materiales = rows.map(row => {
        const sapKey    = Object.keys(row).find(k => k.trim().toUpperCase() === 'SAP');
        const axKey     = Object.keys(row).find(k => k.trim().toUpperCase() === 'AX');
        const nombreKey = Object.keys(row).find(k => k.trim().toLowerCase().includes('nombre'));
        const existKey  = Object.keys(row).find(k => k.trim().toLowerCase().includes('exist'));
        return {
          sapCode: row[sapKey]    != null ? String(row[sapKey]).trim()    : '',
          axCode:  row[axKey]     != null ? String(row[axKey]).trim()     : '',
          name:    row[nombreKey] != null ? String(row[nombreKey]).trim() : '',
          stock:   row[existKey]  != null ? Number(row[existKey]) || 0    : 0,
        };
      }).filter(m => m.name !== '' && (m.sapCode !== '' || m.axCode !== ''));

      if (materiales.length === 0) throw new Error('No se encontraron materiales válidos.');

      const existingSap = new Set(existingItems.map(i => safeStr(i.sapCode,'')).filter(Boolean));
      const existingAx  = new Set(existingItems.map(i => safeStr(i.axCode,'')).filter(Boolean));

      toImport = []; toSkip = [];
      for (const m of materiales) {
        const dupSap = m.sapCode && existingSap.has(m.sapCode);
        const dupAx  = m.axCode  && existingAx.has(m.axCode);
        if (dupSap || dupAx) toSkip.push({ ...m, reason: dupSap ? `SAP ${m.sapCode}` : `AX ${m.axCode}` });
        else toImport.push(m);
      }

      ov.querySelector('#ix-step-file').classList.add('hidden');
      ov.querySelector('#ix-step-preview').classList.remove('hidden');

      ov.querySelector('#ix-summary').innerHTML = [
        { v: toImport.length, l: 'A importar', c: '#2196F3' },
        { v: toSkip.length,   l: 'Saltados',   c: '#E65100' },
        { v: materiales.length, l: 'Total',     c: '#6B7280' },
      ].map(s => `<div class="bg-gray-50 rounded-xl p-2 text-center border border-gray-200">
        <p class="text-xl font-bold" style="color:${s.c}">${s.v}</p>
        <p class="text-xs text-gray-500">${s.l}</p>
      </div>`).join('');

      ov.querySelector('#ix-tabla').innerHTML = `
        <table class="w-full">
          <thead class="bg-gray-50 border-b border-gray-200 sticky top-0">
            <tr>
              <th class="text-left px-3 py-2 font-semibold text-gray-500">Material</th>
              <th class="text-left px-3 py-2 font-semibold text-gray-500">SAP</th>
              <th class="text-right px-3 py-2 font-semibold text-gray-500">Stock</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-gray-100">
            ${toImport.map(m => `<tr>
              <td class="px-3 py-1.5 text-gray-900 truncate max-w-xs">${m.name}</td>
              <td class="px-3 py-1.5 text-gray-400 font-mono">${m.sapCode||'—'}</td>
              <td class="px-3 py-1.5 text-right font-medium text-gray-700">${m.stock}</td>
            </tr>`).join('')}
          </tbody>
        </table>`;

      if (toSkip.length > 0) {
        const sk = ov.querySelector('#ix-skipped');
        sk.classList.remove('hidden');
        sk.innerHTML = `<div class="bg-orange-50 border border-orange-200 rounded-xl p-3">
          <p class="text-xs font-semibold text-orange-800 mb-1">⚠️ ${toSkip.length} saltados (ya existen)</p>
          ${toSkip.slice(0,5).map(m => `<p class="text-xs text-orange-600 truncate">${m.name} — ${m.reason}</p>`).join('')}
          ${toSkip.length > 5 ? `<p class="text-xs text-orange-400">...y ${toSkip.length-5} más</p>` : ''}
        </div>`;
      }
    } catch(e) {
      errEl.textContent = e.message || 'Error al leer el archivo.';
      errEl.classList.remove('hidden');
    }
  });

  ov.querySelector('#ix-back').onclick = () => {
    ov.querySelector('#ix-step-preview').classList.add('hidden');
    ov.querySelector('#ix-step-file').classList.remove('hidden');
    ov.querySelector('#ix-file').value = '';
  };

  ov.querySelector('#ix-submit').onclick = async () => {
    if (toImport.length === 0) return;
    const btn = ov.querySelector('#ix-submit');
    btn.disabled = true; btn.textContent = 'Importando...';
    let ok = 0, errCount = 0;
    for (const m of toImport) {
      try {
        await addDoc(collection(db, 'kardex/inventario/items'), {
          name: m.name, sapCode: m.sapCode, axCode: m.axCode,
          unit: 'unidad', stock: isNaN(m.stock) ? 0 : m.stock,
          minStock: 0, requiereSerial: false,
          creadoEn: serverTimestamp(), creadoPor: session.uid,
        });
        ok++;
      } catch(e) { errCount++; }
      btn.textContent = `Importando ${ok+errCount}/${toImport.length}...`;
    }
    ov.querySelector('#ix-step-preview').classList.add('hidden');
    ov.querySelector('#ix-step-result').classList.remove('hidden');
    ov.querySelector('#ix-result-txt').innerHTML = `
      <p class="text-sm text-gray-700"><strong class="text-green-700">${ok}</strong> material(es) importados.</p>
      ${toSkip.length > 0 ? `<p class="text-sm text-gray-500">${toSkip.length} saltados.</p>` : ''}
      ${errCount > 0 ? `<p class="text-sm text-red-600">${errCount} errores.</p>` : ''}`;
  };

  ov.querySelector('#ix-done').onclick = async () => {
    ov.remove();
    await showInventario(db, session);
  };
}

// ─────────────────────────────────────────
// BORRAR SELECCIONADOS
// ─────────────────────────────────────────
async function showBorrarSeleccionados(db, session, items, onDone) {
  const { deleteDoc, doc: firestoreDoc, getDocs: gd, collection: col } =
    await import('https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js');

  const conMovimientos = [], sinMovimientos = [];
  for (const item of items) {
    const snapTodos = await gd(col(db, 'kardex/movimientos/salidas'));
    const tieneMovs = snapTodos.docs.some(d => (d.data().items||[]).some(i => i.itemId === item.id));
    if (tieneMovs) conMovimientos.push(item);
    else sinMovimientos.push(item);
  }

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Eliminar materiales</h2>
      <button id="bd-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div class="px-5 py-4 space-y-4">
      ${sinMovimientos.length > 0 ? `
      <div>
        <p class="text-sm font-medium text-gray-900 mb-2">Se eliminarán ${sinMovimientos.length} material(es):</p>
        <div class="bg-gray-50 rounded-xl border border-gray-200 divide-y divide-gray-100 max-h-40 overflow-y-auto">
          ${sinMovimientos.map(i => `<div class="px-3 py-2 text-sm text-gray-800">${safeStr(i.name)}</div>`).join('')}
        </div>
      </div>` : ''}
      ${conMovimientos.length > 0 ? `
      <div class="bg-orange-50 border border-orange-200 rounded-xl p-3">
        <p class="text-xs font-semibold text-orange-800 mb-1">⚠️ ${conMovimientos.length} no se puede(n) eliminar — tienen movimientos</p>
        ${conMovimientos.map(i => `<p class="text-xs text-orange-600">• ${safeStr(i.name)}</p>`).join('')}
      </div>` : ''}
      ${sinMovimientos.length === 0 ? `<p class="text-sm text-center text-gray-500 py-4">Ninguno puede eliminarse porque todos tienen movimientos.</p>` : `
      <div class="bg-red-50 border border-red-200 rounded-xl p-3">
        <p class="text-xs text-red-700">Esta acción es <strong>permanente</strong>.</p>
      </div>`}
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="bd-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm">Cancelar</button>
      ${sinMovimientos.length > 0 ? `<button id="bd-confirm" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background:#C62828">Eliminar ${sinMovimientos.length}</button>` : ''}
    </div>`);

  ov.querySelector('#bd-close').onclick = ov.querySelector('#bd-cancel').onclick = () => ov.remove();
  ov.querySelector('#bd-confirm')?.addEventListener('click', async () => {
    const btn = ov.querySelector('#bd-confirm');
    btn.disabled = true; btn.textContent = 'Eliminando...';
    let ok = 0;
    for (const item of sinMovimientos) {
      try { await deleteDoc(firestoreDoc(db, 'kardex/inventario/items', item.id)); ok++; }
      catch(e) { console.error(e); }
    }
    ov.remove();
    showToast(`${ok} material(es) eliminado(s).`, ok > 0 ? 'success' : 'error');
    await onDone();
  });
}

// ─────────────────────────────────────────
// HELPERS UI
// ─────────────────────────────────────────
function mkOverlay(inner) {
  const ov = document.createElement('div');
  ov.className = 'fixed inset-0 bg-black/50 flex items-start justify-center z-50 p-4 overflow-y-auto';
  ov.innerHTML = `<div class="bg-white rounded-2xl shadow-xl w-full max-w-md my-4" style="max-height:92vh;overflow-y:auto;">${inner}</div>`;
  document.body.appendChild(ov);
  return ov;
}

function loading() {
  return `<div class="flex items-center justify-center py-12 gap-3 text-gray-400 text-sm">
    <svg class="animate-spin w-5 h-5" viewBox="0 0 24 24" fill="none">
      <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
      <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
    </svg>Cargando...</div>`;
}

function errHtml() {
  return '<div class="py-12 text-center text-sm text-gray-400">Error al cargar datos.</div>';
}

function fmtDate(ts) {
  if (!ts) return '—';
  try {
    const d = ts.toDate ? ts.toDate() : new Date(ts);
    return d.toLocaleDateString('es-SV', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' });
  } catch { return '—'; }
}

// ─────────────────────────────────────────
// GENERAR PDF VÍA INNOVA-CONVERTER
// Manda JSON → servidor llena Word → convierte PDF → abre
// ─────────────────────────────────────────
const CONVERTER_URL = 'https://innova-converter-558391780384.us-central1.run.app/generar-pdf';

async function generarYAbrirPDF(memo) {
  // Construir payload JSON con los datos del despacho
  const payload = {
    usuarioResponsable:    memo.USUARIO_RESPONSABLE    || '',
    empresaContratista:    memo.EMPRESA_CONTRATISTA    || '',
    instaladorResponsable: memo.INSTALADOR_RESPONSABLE || '',
    entregadoPor:          memo.ENTREGADO_POR          || '',
    placaVehiculo:         memo.PLACA_VEHICULO         || '',
    fechaSolicitud:        memo.FECHA_SOLICITUD        || '',
    fechaEntrega:          memo.FECHA_ENTREGA          || '',
    items: (memo.MATERIALES || []).map(function(i) {
      return {
        sapCode:  String(i.RESERVA  || ''),
        cantidad: String(i.CANTIDAD || ''),
      };
    }),
  };

  const resp = await fetch(CONVERTER_URL, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(payload),
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err.error || 'Error del servidor: ' + resp.status);
  }

  const { pdf_base64 } = await resp.json();
  if (!pdf_base64) throw new Error('El servidor no devolvió el PDF.');

  // Abrir PDF para imprimir
  const pdfBytes = Uint8Array.from(atob(pdf_base64), function(c) { return c.charCodeAt(0); });
  const pdfBlob  = new Blob([pdfBytes], { type: 'application/pdf' });
  const pdfUrl   = URL.createObjectURL(pdfBlob);

  const w = window.open(pdfUrl, '_blank');
  if (!w) {
    const a    = document.createElement('a');
    a.href     = pdfUrl;
    a.download = 'despacho-' + (payload.fechaEntrega || 'sin-fecha') + '.pdf';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
  setTimeout(function() { URL.revokeObjectURL(pdfUrl); }, 10000);
}

// ─────────────────────────────────────────
// STOCK POR USUARIO
// ─────────────────────────────────────────
async function showStockUsuarios(db, session) {
  setTab('usuarios');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();

  try {
    const [snapSalidas, snapItems] = await Promise.all([
      getDocs(query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'))),
      getDocs(collection(db, 'kardex/inventario/items')),
    ]);

    const itemMap = {};
    snapItems.docs.forEach(d => {
      const item = normalizeItem({ id: d.id, ...d.data() });
      if (esValido(item)) itemMap[d.id] = item;
    });

    const stockUsuario = {};
    snapSalidas.docs.forEach(d => {
      const salida  = d.data();
      const usuario = safeStr(salida.usuarioResponsable || salida.tecnicoNombre, '');
      if (!usuario || usuario === '—') return;
      if (!stockUsuario[usuario]) stockUsuario[usuario] = {};
      (salida.items || []).forEach(i => {
        const cant = safeNum(i.cantidad);
        if (!i.itemId || cant <= 0) return;
        stockUsuario[usuario][i.itemId] = (stockUsuario[usuario][i.itemId] || 0) + cant;
      });
    });

    window.__kardexStockUsuario = stockUsuario;

    function getEstado(cant, item) {
      const min = safeNum(item && item.minStock ? item.minStock : 5);
      if (cant <= 0)                      return 'critico';
      if (cant <= Math.max(1, min / 2))   return 'critico';
      if (cant <= min)                    return 'bajo';
      return 'ok';
    }

    const ESTADO_CFG = {
      critico: { icon: '🔴', label: 'Stock crítico', order: 0, bg: '#FEF2F2', border: '#FECACA', color: '#C62828' },
      bajo:    { icon: '🟡', label: 'Stock bajo',    order: 1, bg: '#FFFBEB', border: '#FDE68A', color: '#B45309' },
      ok:      { icon: '🟢', label: 'Suficiente',    order: 2, bg: 'transparent', border: 'transparent', color: '#166534' },
    };

    const usuariosConDatos = USUARIOS_RESPONSABLES.filter(u => stockUsuario[u]);
    if (!usuariosConDatos.length) {
      content.innerHTML = '<div class="bg-white rounded-xl border border-gray-200 py-12 text-center"><p class="text-sm text-gray-400">Sin movimientos registrados aún</p></div>';
      return;
    }

    const expanded = {};
    usuariosConDatos.forEach(u => { expanded[u] = true; });
    let filtro = 'todos';

    function buildBanner(totalCriticos, totalBajos) {
      if (totalCriticos + totalBajos === 0) {
        return '<div class="rounded-xl border px-4 py-3 flex items-center gap-3" style="background:#F0FDF4;border-color:#BBF7D0"><span class="text-xl">✅</span><p class="text-sm font-semibold text-green-800">Todo el material está en niveles normales</p></div>';
      }
      const partes = [];
      if (totalCriticos > 0) partes.push(totalCriticos + ' crítico' + (totalCriticos !== 1 ? 's' : ''));
      if (totalBajos > 0)    partes.push(totalBajos + ' bajo' + (totalBajos !== 1 ? 's' : ''));
      return '<div class="rounded-xl border px-4 py-3 flex items-center gap-3" style="background:#FEF2F2;border-color:#FECACA"><span class="text-xl">⚠️</span><div><p class="text-sm font-bold text-red-800">Atención requerida</p><p class="text-xs text-red-600">' + partes.join(' · ') + '</p></div></div>';
    }

    function buildFiltros() {
      const btns = [
        { key: 'todos',   label: 'Todos' },
        { key: 'critico', label: '🔴 Críticos' },
        { key: 'bajo',    label: '🟡 Bajos' },
      ];
      return '<div class="flex gap-2">' + btns.map(f => {
        const active = filtro === f.key;
        const style  = active ? 'background:#1B4F8A' : '';
        const cls    = active
          ? 'px-3 py-1.5 rounded-xl text-xs font-semibold border border-transparent text-white'
          : 'px-3 py-1.5 rounded-xl text-xs font-semibold border border-gray-200 bg-white text-gray-600';
        return '<button data-filtro="' + f.key + '" class="' + cls + '" style="' + style + '">' + f.label + '</button>';
      }).join('') + '</div>';
    }

    function buildItemRow(e) {
      const est   = ESTADO_CFG[e.estado];
      const name  = safeStr(e.item ? e.item.name : e.id, '—');
      const unit  = safeStr(e.item ? e.item.unit : '', '');
      const sap   = safeStr(e.item ? e.item.sapCode : '', '');
      const min   = safeNum(e.item ? (e.item.minStock || 5) : 5);
      const bg    = e.estado !== 'ok' ? 'background:' + est.bg + ';' : '';
      const sapTxt = sap ? '<p class="text-xs text-gray-400 font-mono">SAP ' + sap + ' · mín. ' + min + ' ' + unit + '</p>' : '';
      return '<div class="flex items-center gap-3 px-4 py-3 border-b border-gray-50 last:border-0" style="' + bg + '">' +
        '<span class="text-base shrink-0">' + est.icon + '</span>' +
        '<div class="flex-1 min-w-0">' +
          '<p class="text-sm font-semibold text-gray-900 truncate leading-tight">' + name + '</p>' +
          '<p class="text-xs mt-0.5 font-medium" style="color:' + est.color + '">' + est.label + '</p>' +
          sapTxt +
        '</div>' +
        '<div class="text-right shrink-0">' +
          '<p class="text-xl font-black" style="color:' + est.color + '">' + e.cant + '</p>' +
          '<p class="text-xs text-gray-400">' + unit + '</p>' +
        '</div>' +
      '</div>';
    }

    function buildUsuarioCard(usuario) {
      const stockU = stockUsuario[usuario] || {};
      let itemsU = Object.entries(stockU)
        .map(function(entry) {
          return { id: entry[0], cant: entry[1], item: itemMap[entry[0]], estado: getEstado(entry[1], itemMap[entry[0]]) };
        })
        .filter(function(e) { return e.cant > 0; });

      itemsU.sort(function(a, b) {
        const od = ESTADO_CFG[a.estado].order - ESTADO_CFG[b.estado].order;
        if (od !== 0) return od;
        return safeStr(a.item ? a.item.name : a.id).localeCompare(safeStr(b.item ? b.item.name : b.id));
      });

      const itemsMostrar = filtro === 'todos' ? itemsU : itemsU.filter(function(e) { return e.estado === filtro; });
      const criticos  = itemsU.filter(function(e) { return e.estado === 'critico'; }).length;
      const bajos     = itemsU.filter(function(e) { return e.estado === 'bajo'; }).length;
      const isOpen    = expanded[usuario];

      const alertaTexto = criticos > 0
        ? '⚠️ ' + criticos + ' crítico' + (criticos !== 1 ? 's' : '')
        : bajos > 0
        ? '⚠️ ' + bajos + ' bajo' + (bajos !== 1 ? 's' : '')
        : '✓ Sin alertas';
      const alertaColor = criticos > 0 ? '#C62828' : bajos > 0 ? '#B45309' : '#166534';
      const chevron = isOpen ? 'rotate-180' : '';

      let detalle = '';
      if (isOpen) {
        if (itemsMostrar.length === 0) {
          detalle = '<p class="text-xs text-gray-400 text-center py-4">Sin materiales en este filtro</p>';
        } else {
          detalle = itemsMostrar.map(buildItemRow).join('');
        }
        detalle = '<div class="border-t border-gray-100">' + detalle + '</div>';
      }

      return '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
        '<button data-toggle="' + usuario + '" class="w-full flex items-center justify-between px-4 py-3.5 text-left">' +
          '<div class="flex items-center gap-3">' +
            '<div class="w-10 h-10 rounded-xl flex items-center justify-center font-bold text-sm text-white shrink-0" style="background:#1B4F8A">' + usuario.slice(0,2) + '</div>' +
            '<div>' +
              '<p class="font-bold text-gray-900 text-base">' + usuario + '</p>' +
              '<p class="text-xs font-semibold" style="color:' + alertaColor + '">' + alertaTexto + '</p>' +
            '</div>' +
          '</div>' +
          '<div class="flex items-center gap-2">' +
            '<span class="text-xs text-gray-400">' + itemsU.length + ' mat.</span>' +
            '<svg class="w-5 h-5 text-gray-400 ' + chevron + ' transition-transform" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>' +
          '</div>' +
        '</button>' +
        detalle +
      '</div>';
    }

    function renderUsuarios() {
      let totalCriticos = 0, totalBajos = 0;
      usuariosConDatos.forEach(function(u) {
        Object.entries(stockUsuario[u] || {}).forEach(function(entry) {
          const est = getEstado(entry[1], itemMap[entry[0]]);
          if (est === 'critico') totalCriticos++;
          else if (est === 'bajo') totalBajos++;
        });
      });

      const html = '<div class="space-y-3">' +
        buildBanner(totalCriticos, totalBajos) +
        buildFiltros() +
        usuariosConDatos.map(buildUsuarioCard).join('') +
        '<p class="text-xs text-gray-400 text-center pb-2">Calculado desde movimientos registrados.</p>' +
      '</div>';

      content.innerHTML = html;

      content.querySelectorAll('[data-toggle]').forEach(function(btn) {
        btn.addEventListener('click', function() {
          expanded[btn.dataset.toggle] = !expanded[btn.dataset.toggle];
          renderUsuarios();
        });
      });
      content.querySelectorAll('[data-filtro]').forEach(function(btn) {
        btn.addEventListener('click', function() {
          filtro = btn.dataset.filtro;
          renderUsuarios();
        });
      });
    }

    renderUsuarios();

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// SOLICITAR MATERIAL (vista campo)
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// CONSUMO DE MATERIAL (campo)
// ─────────────────────────────────────────
const TIPOS_TRABAJO = [
  'Servicio nuevo',
  'Cambio de voltaje',
  'Reconexión',
  'Cambio de medidor',
  'Reubicación de medidor',
  'Reubicación de acometida',
  'Otro',
];

function showRegistrarConsumo(db, session) {
  const usuario = session.usuarioOperativoAsignado;
  if (!usuario) { showToast('Sin usuario operativo asignado.', 'error'); return; }

  // Get items from user stock
  const stockU = window.__kardexStockUsuario?.[usuario] || {};
  const snapItems = window.__kardexItemMap || {};

  // Build material list from user stock
  let misItems = Object.entries(stockU)
    .map(function(e) { return { id: e[0], cant: e[1], item: snapItems[e[0]] }; })
    .filter(function(e) { return e.cant > 0 && e.item; })
    .sort(function(a, b) { return safeStr(a.item.name).localeCompare(safeStr(b.item.name)); });

  let selConsumo = {}; // itemId -> cantidad
  let tipoTrabajo = TIPOS_TRABAJO[0];
  let busqMat = '';

  const ov = mkOverlay(
    '<div class="flex items-center justify-between px-5 py-4 border-b border-gray-100 shrink-0">' +
      '<div><h2 class="font-semibold text-gray-900">Registrar consumo</h2>' +
        '<p class="text-xs text-gray-400 mt-0.5">Material usado en OT · ' + usuario + '</p>' +
      '</div>' +
      '<button id="rc-close" class="text-gray-400 hover:text-gray-700">✕</button>' +
    '</div>' +
    '<div class="flex-1 overflow-y-auto px-5 py-4 space-y-4">' +
      // WO
      '<div>' +
        '<label class="block text-sm font-medium text-gray-700 mb-1.5">Número de OT / WO <span class="text-red-500">*</span></label>' +
        '<input id="rc-wo" type="text" placeholder="Ej. OT-2026-001" class="w-full border border-gray-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>' +
      '</div>' +
      // Tipo de trabajo
      '<div>' +
        '<label class="block text-sm font-medium text-gray-700 mb-1.5">Tipo de trabajo <span class="text-red-500">*</span></label>' +
        '<select id="rc-tipo" class="w-full border border-gray-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400">' +
          TIPOS_TRABAJO.map(function(t) { return '<option value="' + t + '">' + t + '</option>'; }).join('') +
        '</select>' +
      '</div>' +
      '<div id="rc-otro-div" class="hidden">' +
        '<label class="block text-sm font-medium text-gray-700 mb-1.5">Especifica el tipo</label>' +
        '<input id="rc-otro" type="text" placeholder="Tipo de trabajo..." class="w-full border border-gray-200 rounded-xl px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>' +
      '</div>' +
      // Materiales
      '<div>' +
        '<label class="block text-sm font-medium text-gray-700 mb-1.5">Materiales usados <span class="text-red-500">*</span></label>' +
        '<div class="relative mb-2">' +
          '<svg class="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>' +
          '<input id="rc-buscar" type="text" placeholder="Buscar material..." class="w-full border border-gray-200 rounded-xl pl-9 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>' +
        '</div>' +
        '<div id="rc-lista" class="space-y-2 max-h-64 overflow-y-auto"></div>' +
      '</div>' +
      // Resumen seleccionados
      '<div id="rc-resumen" class="hidden">' +
        '<p class="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Material a consumir</p>' +
        '<div id="rc-resumen-items" class="space-y-1.5"></div>' +
      '</div>' +
      '<div id="rc-err" class="hidden text-sm text-red-500 bg-red-50 rounded-xl px-3 py-2"></div>' +
    '</div>' +
    '<div class="px-5 py-4 border-t border-gray-100 flex gap-3 shrink-0">' +
      '<button id="rc-cancel" class="flex-1 border border-gray-200 text-gray-600 font-medium rounded-xl py-3 text-sm">Cancelar</button>' +
      '<button id="rc-submit" class="flex-1 text-white font-medium rounded-xl py-3 text-sm" style="background:#166534">Guardar consumo</button>' +
    '</div>'
  );

  ov.querySelector('#rc-close').onclick = ov.querySelector('#rc-cancel').onclick = () => ov.remove();

  // Tipo trabajo toggle
  ov.querySelector('#rc-tipo').addEventListener('change', function() {
    tipoTrabajo = this.value;
    ov.querySelector('#rc-otro-div').classList.toggle('hidden', this.value !== 'Otro');
  });

  function renderLista() {
    const lista = ov.querySelector('#rc-lista');
    if (!lista) return;
    const q = busqMat.toLowerCase();
    const filtrados = q ? misItems.filter(function(e) {
      return safeStr(e.item.name).toLowerCase().includes(q) ||
             safeStr(e.item.sapCode,'').includes(q);
    }) : misItems;

    if (!filtrados.length) {
      lista.innerHTML = '<p class="text-xs text-gray-400 text-center py-4">' + (q ? 'Sin resultados' : 'Sin material asignado') + '</p>';
      return;
    }

    lista.innerHTML = filtrados.map(function(e) {
      const sel = selConsumo[e.id] || 0;
      const unit = safeStr(e.item.unit, '');
      return '<div class="flex items-center gap-3 bg-gray-50 rounded-xl px-3 py-2.5 border ' + (sel > 0 ? 'border-green-300' : 'border-gray-200') + '">' +
        '<div class="flex-1 min-w-0">' +
          '<p class="text-sm font-medium text-gray-900 leading-tight">' + safeStr(e.item.name) + '</p>' +
          '<p class="text-xs text-gray-400 mt-0.5">Disponible: ' + e.cant + ' ' + unit + '</p>' +
        '</div>' +
        '<div class="flex items-center gap-1.5 shrink-0">' +
          '<button class="rc-dec w-7 h-7 rounded-lg border border-gray-300 bg-white text-gray-600 font-bold flex items-center justify-center active:bg-gray-100" data-iid="' + e.id + '">−</button>' +
          '<span class="w-8 text-center text-sm font-bold ' + (sel > 0 ? 'text-green-700' : 'text-gray-400') + '">' + (sel || 0) + '</span>' +
          '<button class="rc-inc w-7 h-7 rounded-lg border font-bold flex items-center justify-center active:opacity-70 ' + (sel >= e.cant ? 'border-gray-200 text-gray-300' : 'border-blue-300 bg-blue-50 text-blue-600') + '" data-iid="' + e.id + '" data-max="' + e.cant + '">+</button>' +
        '</div>' +
      '</div>';
    }).join('');

    // Wire buttons
    lista.querySelectorAll('.rc-dec').forEach(function(b) {
      b.addEventListener('click', function() {
        const iid = b.dataset.iid;
        if ((selConsumo[iid] || 0) > 0) { selConsumo[iid]--; if (!selConsumo[iid]) delete selConsumo[iid]; }
        renderLista(); renderResumen();
      });
    });
    lista.querySelectorAll('.rc-inc').forEach(function(b) {
      b.addEventListener('click', function() {
        const iid = b.dataset.iid;
        const max = safeNum(b.dataset.max);
        if ((selConsumo[iid] || 0) < max) selConsumo[iid] = (selConsumo[iid] || 0) + 1;
        renderLista(); renderResumen();
      });
    });
  }

  function renderResumen() {
    const resDiv = ov.querySelector('#rc-resumen');
    const resItems = ov.querySelector('#rc-resumen-items');
    const entries = Object.entries(selConsumo).filter(function(e) { return e[1] > 0; });
    if (!entries.length) { resDiv.classList.add('hidden'); return; }
    resDiv.classList.remove('hidden');
    resItems.innerHTML = entries.map(function(e) {
      const itData = misItems.find(function(m) { return m.id === e[0]; });
      if (!itData) return '';
      return '<div class="flex items-center justify-between px-3 py-2 bg-green-50 rounded-lg border border-green-200">' +
        '<p class="text-sm text-gray-800 flex-1 leading-tight">' + safeStr(itData.item.name) + '</p>' +
        '<span class="text-sm font-bold text-green-700 ml-2">' + e[1] + ' ' + safeStr(itData.item.unit,'') + '</span>' +
      '</div>';
    }).join('');
  }

  ov.querySelector('#rc-buscar').addEventListener('input', function(e) {
    busqMat = e.target.value.trim();
    renderLista();
  });

  renderLista();

  ov.querySelector('#rc-submit').onclick = async function() {
    const errEl  = ov.querySelector('#rc-err');
    const btn    = ov.querySelector('#rc-submit');
    const wo     = ov.querySelector('#rc-wo').value.trim();
    const tipo   = ov.querySelector('#rc-tipo').value;
    const otro   = ov.querySelector('#rc-otro')?.value.trim() || '';
    const tipoFinal = tipo === 'Otro' ? (otro || 'Otro') : tipo;
    errEl.classList.add('hidden');

    if (!wo) { errEl.textContent = 'Ingresa el número de OT/WO.'; errEl.classList.remove('hidden'); return; }
    const items = Object.entries(selConsumo).filter(function(e) { return e[1] > 0; });
    if (!items.length) { errEl.textContent = 'Selecciona al menos un material.'; errEl.classList.remove('hidden'); return; }

    btn.disabled = true; btn.textContent = 'Guardando...';
    try {
      const consumoItems = items.map(function(e) {
        const itData = misItems.find(function(m) { return m.id === e[0]; });
        return { itemId: e[0], nombre: safeStr(itData?.item.name), unit: safeStr(itData?.item.unit,''), sapCode: safeStr(itData?.item.sapCode,''), cantidad: e[1] };
      });

      await addDoc(collection(db, 'kardex/consumos'), {
        wo: wo,
        tipoTrabajo: tipoFinal,
        usuarioOperativo: usuario,
        registradoPor: session.uid,
        registradoPorNombre: session.displayName,
        items: consumoItems,
        fecha: serverTimestamp(),
      });

      // Discount from user stock cache
      if (window.__kardexStockUsuario) {
        if (!window.__kardexStockUsuario[usuario]) window.__kardexStockUsuario[usuario] = {};
        items.forEach(function(e) {
          window.__kardexStockUsuario[usuario][e[0]] = Math.max(0, (window.__kardexStockUsuario[usuario][e[0]] || 0) - e[1]);
        });
      }

      ov.remove();
      showToast('Consumo registrado correctamente.', 'success');
      await showDashboard(db, session);
    } catch(e) {
      errEl.textContent = 'Error al guardar. Intenta de nuevo.';
      errEl.classList.remove('hidden');
      btn.disabled = false; btn.textContent = 'Guardar consumo';
      console.error(e);
    }
  };
}

// ─────────────────────────────────────────
// HISTORIAL DE CONSUMOS (campo)
// ─────────────────────────────────────────
async function showMisConsumos(db, session) {
  setTab('consumo');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  const usuario = session.usuarioOperativoAsignado;

  try {
    const snap = await getDocs(query(
      collection(db, 'kardex/consumos'),
      where('usuarioOperativo', '==', usuario)
    ));
    const consumos = snap.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });
    consumos.sort(function(a, b) { return (b.fecha?.seconds||0) - (a.fecha?.seconds||0); });

    content.innerHTML =
      '<div class="space-y-3">' +
        '<div class="flex items-center justify-between">' +
          '<p class="font-semibold text-gray-900">Mis consumos</p>' +
          '<button id="btn-nuevo-consumo" class="text-xs font-medium px-3 py-1.5 rounded-lg text-white" style="background:#166534">+ Registrar</button>' +
        '</div>' +
        (consumos.length === 0
          ? '<div class="bg-white rounded-xl border border-gray-200 py-12 text-center"><p class="text-sm text-gray-400">Sin consumos registrados</p></div>'
          : '<div class="space-y-2">' +
              consumos.map(function(c) {
                return '<div class="bg-white rounded-xl border border-gray-200 px-4 py-3">' +
                  '<div class="flex items-start justify-between gap-2 mb-2">' +
                    '<div>' +
                      '<p class="text-sm font-bold text-gray-900">' + safeStr(c.wo) + '</p>' +
                      '<p class="text-xs text-gray-400 mt-0.5">' + fmtDate(c.fecha) + ' · ' + safeStr(c.tipoTrabajo) + '</p>' +
                    '</div>' +
                    '<span class="text-xs font-semibold px-2 py-0.5 rounded-full shrink-0" style="background:#DCFCE7;color:#166534">✓ Registrado</span>' +
                  '</div>' +
                  '<div class="flex flex-wrap gap-1">' +
                    (c.items||[]).map(function(i) {
                      return '<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">' + i.cantidad + ' ' + safeStr(i.unit,'') + ' ' + safeStr(i.nombre) + '</span>';
                    }).join('') +
                  '</div>' +
                '</div>';
              }).join('') +
            '</div>'
        ) +
      '</div>';

    content.querySelector('#btn-nuevo-consumo')?.addEventListener('click', function() {
      showRegistrarConsumo(db, session);
    });

    // Rebind tabs
    document.querySelectorAll('.ktab').forEach(function(b) {
      b.addEventListener('click', async function() {
        const t = b.dataset.tab;
        if (t === 'dashboard')   await showDashboard(db, session);
        if (t === 'consumo')     await showMisConsumos(db, session);
        if (t === 'solicitar')   await showSolicitarMaterial(db, session);
        if (t === 'mis-pedidos') await showMisSolicitudes(db, session);
      });
    });

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

async function showSolicitarMaterial(db, session) {
  setTab('solicitar');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();

  try {
    const snapItems = await getDocs(collection(db, 'kardex/inventario/items'));
    const items = snapItems.docs
      .map(d => normalizeItem({ id: d.id, ...d.data() }))
      .filter(i => esValido(i) && safeNum(i.stock) > 0)
      .sort((a,b) => safeStr(a.name).localeCompare(safeStr(b.name)));

    let sel  = [];   // { itemId, name, unit, stock, cantidad }
    let busq = '';

    function tc(str) {
      return safeStr(str).toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    }

    function render() {
      const html =
        '<div class="space-y-4">' +

          // Carrito
          '<div>' +
            '<div class="flex items-center justify-between mb-2">' +
              '<p class="text-sm font-semibold text-gray-700">Tu pedido</p>' +
              (sel.length > 0 ? '<span class="text-xs font-bold px-2 py-0.5 rounded-full text-white" style="background:#1B4F8A">' + sel.length + '</span>' : '') +
            '</div>' +
            (sel.length === 0
              ? '<div class="border-2 border-dashed border-gray-200 rounded-2xl py-6 text-center"><p class="text-sm text-gray-400">Busca materiales abajo y tócalos para agregar</p></div>'
              : '<div class="space-y-2">' + sel.map(function(s, idx) {
                  return '<div class="bg-white rounded-2xl border-2 border-blue-200 px-4 py-3 flex items-center gap-3">' +
                    '<div class="flex-1 min-w-0">' +
                      '<p class="text-sm font-semibold text-gray-900 truncate">' + tc(s.name) + '</p>' +
                      '<p class="text-xs text-gray-400">' + s.stock + ' ' + safeStr(s.unit,'') + ' disponibles</p>' +
                    '</div>' +
                    '<div class="flex items-center gap-1.5 shrink-0">' +
                      '<button data-dec="' + idx + '" class="w-8 h-8 rounded-xl border border-gray-300 bg-gray-50 text-lg font-bold text-gray-500 flex items-center justify-center">−</button>' +
                      '<span class="w-8 text-center text-base font-bold text-gray-900">' + s.cantidad + '</span>' +
                      '<button data-inc="' + idx + '" class="w-8 h-8 rounded-xl border border-blue-300 bg-blue-50 text-lg font-bold text-blue-600 flex items-center justify-center ' + (s.cantidad >= s.stock ? 'opacity-40 cursor-not-allowed' : '') + '" ' + (s.cantidad >= s.stock ? 'disabled' : '') + '>+</button>' +
                    '</div>' +
                    '<button data-del="' + idx + '" class="w-7 h-7 rounded-xl bg-red-50 flex items-center justify-center shrink-0"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#EF4444" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>' +
                  '</div>';
                }).join('') + '</div>'
            ) +
          '</div>' +

          // Buscador
          '<div>' +
            '<p class="text-sm font-semibold text-gray-700 mb-2">Materiales disponibles</p>' +
            '<div class="relative mb-2">' +
              '<svg class="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>' +
              '<input id="sol-buscar" type="text" value="' + busq + '" placeholder="Buscar material..." autocomplete="off" class="w-full bg-white border-2 border-gray-200 rounded-2xl pl-10 pr-4 py-2.5 text-sm focus:outline-none focus:border-blue-400"/>' +
            '</div>' +
            '<div id="sol-lista" class="space-y-1.5"></div>' +
          '</div>' +

          // Error
          '<div id="sol-err" class="hidden text-sm text-red-600 text-center"></div>' +

          // Botón
          '<button id="sol-submit" class="w-full font-bold rounded-2xl py-4 text-sm text-white transition-all" style="background:' + (sel.length > 0 ? '#1B4F8A' : '#D1D5DB') + '" ' + (sel.length === 0 ? 'disabled' : '') + '>' +
            (sel.length > 0 ? 'Enviar solicitud · ' + sel.length + ' material' + (sel.length !== 1 ? 'es' : '') : 'Selecciona materiales para continuar') +
          '</button>' +

        '</div>';

      content.innerHTML = html;

      // Carrito events
      content.querySelectorAll('[data-dec]').forEach(function(b) {
        b.onclick = function() { if (sel[+b.dataset.dec].cantidad > 1) { sel[+b.dataset.dec].cantidad--; render(); } };
      });
      content.querySelectorAll('[data-inc]').forEach(function(b) {
        b.onclick = function() { if (sel[+b.dataset.inc].cantidad < sel[+b.dataset.inc].stock) { sel[+b.dataset.inc].cantidad++; render(); } };
      });
      content.querySelectorAll('[data-del]').forEach(function(b) {
        b.onclick = function() { sel.splice(+b.dataset.del, 1); render(); };
      });

      // Buscador
      content.querySelector('#sol-buscar').addEventListener('input', function(e) {
        busq = e.target.value;
        renderLista();
      });

      // Submit
      content.querySelector('#sol-submit').addEventListener('click', handleEnviar);

      renderLista();
    }

    function renderLista() {
      const el = content.querySelector('#sol-lista');
      if (!el) return;
      const selIds = new Set(sel.map(function(s) { return s.itemId; }));
      const q = busq.trim().toLowerCase();
      const lista = q
        ? items.filter(function(i) { return safeStr(i.name).toLowerCase().includes(q); })
        : items;

      if (!lista.length) {
        el.innerHTML = '<p class="text-sm text-gray-400 text-center py-4">Sin resultados</p>';
        return;
      }

      el.innerHTML = lista.map(function(item) {
        const agregado = selIds.has(item.id);
        const stockColor = safeNum(item.stock) <= safeNum(item.minStock || 5) ? '#E65100' : '#374151';
        return '<button data-item="' + item.id + '" ' + (agregado ? 'disabled' : '') + ' class="w-full flex items-center justify-between gap-3 text-left px-4 py-3 rounded-2xl border-2 ' + (agregado ? 'border-green-200 bg-green-50 cursor-not-allowed' : 'border-gray-200 bg-white active:border-blue-400 active:bg-blue-50') + '">' +
          '<p class="text-sm font-semibold truncate ' + (agregado ? 'text-green-700' : 'text-gray-900') + '">' + tc(item.name) + '</p>' +
          (agregado
            ? '<span class="text-xs font-bold text-green-600 shrink-0">✓</span>'
            : '<span class="text-sm font-bold shrink-0" style="color:' + stockColor + '">' + item.stock + ' <span class="text-xs font-normal text-gray-400">' + safeStr(item.unit,'') + '</span></span>'
          ) +
        '</button>';
      }).join('');

      el.querySelectorAll('[data-item]').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const item = items.find(function(i) { return i.id === btn.dataset.item; });
          if (!item || sel.some(function(s) { return s.itemId === item.id; })) return;
          showModalCantidadSol(item);
        });
      });
    }

    function showModalCantidadSol(item) {
      const modal = document.createElement('div');
      modal.className = 'fixed inset-0 bg-black/50 flex items-end z-50';
      modal.innerHTML =
        '<div class="bg-white w-full rounded-t-3xl px-5 pt-5 pb-8 space-y-4">' +
          '<div class="w-10 h-1 bg-gray-200 rounded-full mx-auto"></div>' +
          '<div>' +
            '<p class="font-bold text-gray-900 text-lg">' + tc(item.name) + '</p>' +
            '<p class="text-sm text-gray-400">' + item.stock + ' ' + safeStr(item.unit,'') + ' disponibles</p>' +
          '</div>' +
          '<div class="flex items-center justify-between gap-4">' +
            '<button id="mcs-dec" class="w-16 h-16 rounded-2xl border-2 border-gray-200 bg-gray-50 text-3xl font-bold text-gray-400 flex items-center justify-center">−</button>' +
            '<div class="flex-1 text-center">' +
              '<input id="mcs-cant" type="number" min="1" max="' + item.stock + '" value="1" class="w-full text-center text-5xl font-black text-gray-900 bg-transparent border-none focus:outline-none"/>' +
              '<p class="text-sm text-gray-400">' + safeStr(item.unit,'') + '</p>' +
            '</div>' +
            '<button id="mcs-inc" class="w-16 h-16 rounded-2xl border-2 border-blue-300 bg-blue-50 text-3xl font-bold text-blue-600 flex items-center justify-center">+</button>' +
          '</div>' +
          '<div id="mcs-err" class="hidden text-sm text-red-500 text-center"></div>' +
          '<button id="mcs-add" class="w-full text-white font-bold rounded-2xl py-4" style="background:#1B4F8A">Agregar al pedido</button>' +
        '</div>';
      document.body.appendChild(modal);

      const cantEl = modal.querySelector('#mcs-cant');
      setTimeout(function() { cantEl.focus(); cantEl.select(); }, 80);

      modal.addEventListener('click', function(e) { if (e.target === modal) modal.remove(); });
      modal.querySelector('#mcs-dec').onclick = function() { const v=safeNum(cantEl.value); if(v>1) cantEl.value=v-1; };
      modal.querySelector('#mcs-inc').onclick = function() { const v=safeNum(cantEl.value); if(v<item.stock) cantEl.value=v+1; };

      modal.querySelector('#mcs-add').addEventListener('click', function() {
        const cant  = safeNum(cantEl.value);
        const errEl = modal.querySelector('#mcs-err');
        if (cant <= 0)         { errEl.textContent='Ingresa una cantidad mayor a 0.'; errEl.classList.remove('hidden'); return; }
        if (cant > item.stock) { errEl.textContent='Máximo: ' + item.stock + ' ' + safeStr(item.unit,''); errEl.classList.remove('hidden'); return; }
        sel.push({ itemId:item.id, name:item.name, unit:item.unit, stock:item.stock, cantidad:cant });
        modal.remove();
        render();
      });
    }

    async function handleEnviar() {
      if (!sel.length) return;
      const btn   = content.querySelector('#sol-submit');
      const errEl = content.querySelector('#sol-err');
      errEl.classList.add('hidden');
      btn.disabled = true;
      btn.textContent = 'Enviando...';

      try {
        await addDoc(collection(db, 'solicitudes_material'), {
          usuarioUid:              session.uid,
          usuarioNombre:           session.displayName,
          usuarioRole:             session.role,
          usuarioOperativo:        session.usuarioOperativoAsignado || null,
          materiales: sel.map(function(s) {
            return { itemId:s.itemId, nombre:s.name, unit:s.unit, cantidad:s.cantidad };
          }),
          estado:    'pendiente',
          fecha:     serverTimestamp(),
          notas:     '',
        });
        showToast('Solicitud enviada correctamente.', 'success');
        sel = [];
        render();
      } catch(e) {
        errEl.textContent = 'Error al enviar. Intenta de nuevo.';
        errEl.classList.remove('hidden');
        btn.disabled = false;
        btn.textContent = 'Enviar solicitud · ' + sel.length + ' material' + (sel.length !== 1 ? 'es' : '');
        console.error(e);
      }
    }

    render();

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// MIS SOLICITUDES (vista campo)
// ─────────────────────────────────────────
async function showMisSolicitudes(db, session) {
  setTab('mis-pedidos');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();

  try {
    const snap = await getDocs(query(
      collection(db, 'solicitudes_material'),
      where('usuarioUid', '==', session.uid)
    ));
    const solicitudes = snap.docs
      .map(function(d) { return Object.assign({ id: d.id }, d.data()); })
      .sort(function(a, b) { return (b.fecha?.seconds||0) - (a.fecha?.seconds||0); });

    const ESTADO_BADGE = {
      pendiente:  { label: 'Pendiente',  bg: '#FEF3C7', color: '#B45309' },
      aprobado:   { label: 'Aprobado',   bg: '#DCFCE7', color: '#166534' },
      rechazado:  { label: 'Rechazado',  bg: '#FEE2E2', color: '#C62828' },
    };

    function tc(str) {
      return safeStr(str).toLowerCase().replace(/\b\w/g, function(c) { return c.toUpperCase(); });
    }

    if (!solicitudes.length) {
      content.innerHTML =
        '<div class="bg-white rounded-xl border border-gray-200 py-12 text-center space-y-2">' +
          '<p class="text-2xl">📦</p>' +
          '<p class="text-sm font-medium text-gray-600">Sin solicitudes aún</p>' +
          '<p class="text-xs text-gray-400">Usa la pestaña "Solicitar" para pedir materiales</p>' +
        '</div>';
      return;
    }

    const rows = solicitudes.map(function(s) {
      const badge = ESTADO_BADGE[s.estado] || ESTADO_BADGE.pendiente;
      const mats  = (s.materiales || []).map(function(m) {
        return '<span class="inline-flex items-center gap-1 text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">' + m.cantidad + ' ' + tc(m.nombre) + '</span>';
      }).join(' ');
      return '<div class="px-4 py-3 border-b border-gray-100 last:border-0">' +
        '<div class="flex items-center justify-between mb-1.5">' +
          '<p class="text-xs text-gray-400">' + fmtDate(s.fecha) + '</p>' +
          '<span class="text-xs font-semibold px-2 py-0.5 rounded-full" style="background:' + badge.bg + ';color:' + badge.color + '">' + badge.label + '</span>' +
        '</div>' +
        '<div class="flex flex-wrap gap-1">' + mats + '</div>' +
        (s.notas ? '<p class="text-xs text-gray-500 mt-1 italic">' + s.notas + '</p>' : '') +
      '</div>';
    }).join('');

    content.innerHTML =
      '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
        '<div class="px-4 py-3 border-b border-gray-100">' +
          '<p class="font-semibold text-sm text-gray-900">Mis solicitudes</p>' +
          '<p class="text-xs text-gray-400 mt-0.5">' + solicitudes.length + ' registro(s)</p>' +
        '</div>' +
        rows +
      '</div>';

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// GESTIÓN DE SOLICITUDES (vista admin/coordinadora)
// Flujo: revisar → ajustar → aprobar → despacho prellenado
// ─────────────────────────────────────────
async function showSolicitudes(db, session) {
  setTab('solicitudes');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();

  try {
    const snap = await getDocs(query(
      collection(db, 'solicitudes_material'),
      orderBy('fecha', 'desc')
    ));
    const solicitudes = snap.docs.map(function(d) { return Object.assign({ id: d.id }, d.data()); });
    const pendientes  = solicitudes.filter(function(s) { return s.estado === 'pendiente'; });

    const ESTADO_BADGE = {
      pendiente:           { label: 'Pendiente',            bg: '#FEF3C7', color: '#B45309' },
      aprobado:            { label: 'Aprobado',             bg: '#DCFCE7', color: '#166534' },
      rechazado:           { label: 'Rechazado',            bg: '#FEE2E2', color: '#C62828' },
      listo_para_despacho: { label: 'Listo para despacho',  bg: '#EFF6FF', color: '#1D4ED8' },
    };

    function tc(str) {
      return safeStr(str).toLowerCase().replace(/\w/g, function(c) { return c.toUpperCase(); });
    }

    function buildRows() {
      return solicitudes.map(function(s) {
        const badge = ESTADO_BADGE[s.estado] || ESTADO_BADGE.pendiente;
        const mats = (s.materiales || []).map(function(m) {
          return '<div class="flex items-center justify-between py-1.5 border-b border-gray-50 last:border-0">' +
            '<p class="text-sm text-gray-800">' + tc(m.nombre) + '</p>' +
            '<p class="text-sm font-bold text-gray-900">' + m.cantidad + ' <span class="text-xs font-normal text-gray-400">' + safeStr(m.unit,'') + '</span></p>' +
          '</div>';
        }).join('');

        let acciones = '';
        if (s.estado === 'pendiente') {
          acciones =
            '<div class="flex gap-2 mt-3">' +
              '<button data-rechazar="' + s.id + '" class="flex-1 border border-red-200 text-red-600 font-medium rounded-xl py-2.5 text-sm bg-red-50">Rechazar</button>' +
              '<button data-revisar="' + s.id + '" class="flex-1 text-white font-semibold rounded-xl py-2.5 text-sm" style="background:#1B4F8A">Revisar y aprobar →</button>' +
            '</div>';
        } else if (s.estado === 'listo_para_despacho') {
          acciones =
            '<div class="mt-3">' +
              '<button data-despachar="' + s.id + '" class="w-full text-white font-semibold rounded-xl py-3 text-sm" style="background:#166534">📦 Completar despacho</button>' +
            '</div>';
        }

        return '<div class="px-4 py-4 border-b border-gray-100 last:border-0">' +
          '<div class="flex items-center justify-between mb-2">' +
            '<div>' +
              '<p class="font-bold text-gray-900">' + safeStr(s.usuarioNombre) + '</p>' +
              '<p class="text-xs text-gray-400">' + fmtDate(s.fecha) + (s.usuarioOperativo ? ' · ' + s.usuarioOperativo : '') + '</p>' +
            '</div>' +
            '<span class="text-xs font-semibold px-2 py-1 rounded-full" style="background:' + badge.bg + ';color:' + badge.color + '">' + badge.label + '</span>' +
          '</div>' +
          '<div class="border border-gray-100 rounded-xl px-3 py-1 bg-gray-50">' + mats + '</div>' +
          (s.notas ? '<p class="text-xs text-gray-500 mt-2 italic">' + s.notas + '</p>' : '') +
          acciones +
        '</div>';
      }).join('');
    }

    function render() {
      const pendientesBadge = pendientes.length > 0
        ? ' <span class="ml-1 text-xs font-bold px-1.5 py-0.5 rounded-full text-white" style="background:#C62828">' + pendientes.length + '</span>'
        : '';

      content.innerHTML =
        '<div class="space-y-3">' +
          '<div class="bg-white rounded-xl border border-gray-200 overflow-hidden">' +
            '<div class="px-4 py-3 border-b border-gray-100">' +
              '<p class="font-semibold text-sm text-gray-900">Solicitudes de material' + pendientesBadge + '</p>' +
            '</div>' +
            (solicitudes.length === 0
              ? '<p class="text-sm text-gray-400 text-center py-10">Sin solicitudes aún</p>'
              : buildRows()
            ) +
          '</div>' +
        '</div>';

      // Rechazar
      content.querySelectorAll('[data-rechazar]').forEach(function(btn) {
        btn.addEventListener('click', async function() {
          btn.disabled = true; btn.textContent = 'Rechazando...';
          try {
            await updateDoc(doc(db, 'solicitudes_material', btn.dataset.rechazar), {
              estado: 'rechazado',
              rechazadoPor: session.uid,
              rechazadoPorNombre: session.displayName,
              fechaRespuesta: serverTimestamp(),
            });
            showToast('Solicitud rechazada.', 'success');
            await showSolicitudes(db, session);
          } catch(e) { showToast('Error.', 'error'); btn.disabled = false; btn.textContent = 'Rechazar'; }
        });
      });

      // Revisar — abre modal para ajustar y aprobar
      content.querySelectorAll('[data-revisar]').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const sol = solicitudes.find(function(s) { return s.id === btn.dataset.revisar; });
          if (sol) showModalRevisar(db, session, sol);
        });
      });

      // Completar despacho — abre formulario prellenado
      content.querySelectorAll('[data-despachar]').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const sol = solicitudes.find(function(s) { return s.id === btn.dataset.despachar; });
          if (sol) abrirDespachoDesdeSolicitud(db, session, sol);
        });
      });
    }

    render();

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// MODAL REVISAR SOLICITUD
// Permite ajustar cantidades antes de aprobar
// ─────────────────────────────────────────
function showModalRevisar(db, session, sol) {
  // Copia editable de materiales
  let mats = (sol.materiales || []).map(function(m) {
    return Object.assign({}, m);
  });

  function tc(str) {
    return safeStr(str).toLowerCase().replace(/\w/g, function(c) { return c.toUpperCase(); });
  }

  const modal = document.createElement('div');
  modal.className = 'fixed inset-0 bg-black/50 flex items-end z-50';
  document.body.appendChild(modal);

  function renderModal() {
    modal.innerHTML =
      '<div class="bg-white w-full rounded-t-3xl" style="max-height:90dvh;overflow-y:auto">' +
        '<div class="px-5 pt-5 pb-2">' +
          '<div class="w-10 h-1 bg-gray-200 rounded-full mx-auto mb-4"></div>' +
          '<div class="flex items-center justify-between mb-1">' +
            '<h2 class="font-bold text-gray-900 text-lg">Revisar solicitud</h2>' +
            '<button id="mr-close" class="text-gray-400 text-2xl leading-none">✕</button>' +
          '</div>' +
          '<p class="text-sm text-gray-500 mb-4">' + safeStr(sol.usuarioNombre) + (sol.usuarioOperativo ? ' · ' + sol.usuarioOperativo : '') + '</p>' +

          // Materiales editables
          '<p class="text-xs font-semibold text-gray-400 uppercase tracking-wider mb-2">Materiales solicitados</p>' +
          '<div class="space-y-2 mb-4">' +
          mats.map(function(m, idx) {
            return '<div class="flex items-center gap-3 bg-gray-50 rounded-xl px-3 py-2.5">' +
              '<div class="flex-1 min-w-0">' +
                '<p class="text-sm font-semibold text-gray-900 truncate">' + tc(m.nombre) + '</p>' +
                '<p class="text-xs text-gray-400">' + safeStr(m.unit,'') + '</p>' +
              '</div>' +
              '<div class="flex items-center gap-1.5 shrink-0">' +
                '<button data-mdec="' + idx + '" class="w-8 h-8 rounded-lg border border-gray-300 bg-white text-lg font-bold text-gray-600 flex items-center justify-center">−</button>' +
                '<span class="w-10 text-center text-base font-bold text-gray-900">' + m.cantidad + '</span>' +
                '<button data-minc="' + idx + '" class="w-8 h-8 rounded-lg border border-blue-300 bg-blue-50 text-lg font-bold text-blue-600 flex items-center justify-center">+</button>' +
              '</div>' +
              '<button data-mrem="' + idx + '" class="w-7 h-7 rounded-lg bg-red-50 flex items-center justify-center shrink-0 text-red-400 text-sm">✕</button>' +
            '</div>';
          }).join('') +
          '</div>' +

          // Botones
          '<div class="flex gap-3 pb-6">' +
            '<button id="mr-rechazar" class="flex-1 border border-red-200 text-red-600 font-medium rounded-2xl py-3 text-sm bg-red-50">Rechazar</button>' +
            '<button id="mr-aprobar" class="flex-1 text-white font-bold rounded-2xl py-3 text-sm" style="background:#166534">✓ Aprobar</button>' +
          '</div>' +
        '</div>' +
      '</div>';

    modal.querySelector('#mr-close').onclick = function() { modal.remove(); };

    modal.querySelectorAll('[data-mdec]').forEach(function(b) {
      b.onclick = function() {
        const idx = parseInt(b.dataset.mdec);
        if (mats[idx].cantidad > 1) { mats[idx].cantidad--; renderModal(); }
      };
    });
    modal.querySelectorAll('[data-minc]').forEach(function(b) {
      b.onclick = function() { mats[parseInt(b.dataset.minc)].cantidad++; renderModal(); };
    });
    modal.querySelectorAll('[data-mrem]').forEach(function(b) {
      b.onclick = function() { mats.splice(parseInt(b.dataset.mrem), 1); renderModal(); };
    });

    modal.querySelector('#mr-rechazar').addEventListener('click', async function() {
      const btn = modal.querySelector('#mr-rechazar');
      btn.disabled = true; btn.textContent = 'Rechazando...';
      try {
        await updateDoc(doc(db, 'solicitudes_material', sol.id), {
          estado: 'rechazado',
          rechazadoPor: session.uid,
          rechazadoPorNombre: session.displayName,
          fechaRespuesta: serverTimestamp(),
        });
        modal.remove();
        showToast('Solicitud rechazada.', 'success');
        await showSolicitudes(db, session);
      } catch(e) { showToast('Error.', 'error'); btn.disabled = false; btn.textContent = 'Rechazar'; }
    });

    modal.querySelector('#mr-aprobar').addEventListener('click', async function() {
      if (mats.length === 0) { showToast('Agrega al menos un material.', 'error'); return; }
      const btn = modal.querySelector('#mr-aprobar');
      btn.disabled = true; btn.textContent = 'Aprobando...';
      try {
        // Guardar materiales aprobados y estado
        await updateDoc(doc(db, 'solicitudes_material', sol.id), {
          estado: 'listo_para_despacho',
          materialesAprobados: mats,
          aprobadoPor: session.uid,
          aprobadoPorNombre: session.displayName,
          fechaAprobacion: serverTimestamp(),
        });
        // Cerrar modal y abrir INMEDIATAMENTE el formulario de despacho
        modal.remove();
        const solActualizada = Object.assign({}, sol, {
          estado: 'listo_para_despacho',
          materialesAprobados: mats,
        });
        await abrirDespachoDesdeSolicitud(db, session, solActualizada);
      } catch(e) {
        showToast('Error al aprobar.', 'error');
        btn.disabled = false; btn.textContent = '✓ Aprobar';
        console.error(e);
      }
    });
  }

  renderModal();
  modal.addEventListener('click', function(e) { if (e.target === modal) modal.remove(); });
}

// ─────────────────────────────────────────
// ABRIR DESPACHO DESDE SOLICITUD
// Prellenar formulario con materiales aprobados
// ─────────────────────────────────────────
async function abrirDespachoDesdeSolicitud(db, session, sol) {
  // Cargar inventario para verificar stock
  const snapI = await getDocs(collection(db, 'kardex/inventario/items'));
  const itemsMap = {};
  snapI.docs.forEach(function(d) {
    const item = normalizeItem({ id: d.id, ...d.data() });
    if (esValido(item)) itemsMap[d.id] = item;
  });

  // Materiales aprobados de la solicitud
  const matsAprobados = sol.materialesAprobados || sol.materiales || [];

  // Buscar items del inventario por nombre (match por nombre normalizado)
  const sel = matsAprobados.map(function(m) {
    // Intentar encontrar el item en inventario
    const found = Object.values(itemsMap).find(function(i) {
      return safeStr(i.name).toLowerCase() === safeStr(m.nombre).toLowerCase() ||
             i.id === m.itemId;
    });
    if (!found) return null;
    return {
      itemId: found.id, name: found.name, unit: found.unit,
      sapCode: found.sapCode, axCode: found.axCode,
      stock: found.stock, cantidad: safeNum(m.cantidad),
      requiereSerial: found.requiereSerial,
      modoSerial: 'individual', seriales: [], serialInicio: '', serialFin: '',
    };
  }).filter(Boolean);

  // Pre-rellenar usuario responsable con el usuario operativo de la solicitud
  const hdrPreset = {
    responsable: sol.usuarioOperativo || '',
    contratista: EMPRESAS_CONTRATISTAS[0] || 'INNOVA',
    instalador: '',
    placaSel: '', placaOtro: '',
    fechaSol: today(), fechaEnt: today(),
    _solicitudId: sol.id, // referencia para marcar como despachada
  };

  // Llamar al formulario de salida con datos prellenados
  showFormSalidaConPreset(db, session, sel, hdrPreset);
}

// ─────────────────────────────────────────
// FORMULARIO DE SALIDA CON PRESET
// Igual que showFormSalida pero con datos precargados
// ─────────────────────────────────────────
async function showFormSalidaConPreset(db, session, selInicial, hdrPreset) {
  const [snapI, snapU] = await Promise.all([
    getDocs(collection(db, 'kardex/inventario/items')),
    getDocs(query(collection(db, 'users'), where('active','==',true))),
  ]);

  const items = snapI.docs
    .map(d => normalizeItem({ id: d.id, ...d.data() }))
    .filter(i => esValido(i) && safeNum(i.stock) > 0)
    .sort((a,b) => safeStr(a.name).localeCompare(safeStr(b.name)));

  const usuarios = snapU.docs
    .map(d => d.data())
    .filter(u => ['campo','coordinadora'].includes(u.role))
    .sort((a,b) => safeStr(a.displayName).localeCompare(safeStr(b.displayName)));

  let sel = selInicial || [];
  let step = 1;
  let hdr = hdrPreset || {
    responsable: '', contratista: EMPRESAS_CONTRATISTAS[0]||'INNOVA',
    instalador: '', placaSel: '', placaOtro: '',
    fechaSol: today(), fechaEnt: today(),
  };
  let busq = '';

  const solicitudId = hdr._solicitudId || null;

  // Reusar exactamente el mismo código de showFormSalida
  // pero pasando sel y hdr precargados
  const ov = document.createElement('div');
  ov.className = 'fixed inset-0 z-50';
  document.body.appendChild(ov);

  function tc(str) {
    return safeStr(str).toLowerCase().replace(/\w/g, c => c.toUpperCase());
  }
  function getRec() {
    try { return JSON.parse(sessionStorage.getItem('kardex_rec')||'[]'); } catch { return []; }
  }
  function addRec(id) {
    const p = getRec().filter(x=>x!==id);
    sessionStorage.setItem('kardex_rec', JSON.stringify([id,...p].slice(0,6)));
  }

  // Siempre empezar en paso 1 para llenar encabezado obligatorio
  step = 1;

  function renderStep1() {
    ov.innerHTML = '<div class="flex flex-col h-full bg-white" style="max-height:100dvh">' +
      '<div class="flex items-center gap-3 px-4 py-3 border-b border-gray-100 shrink-0">' +
        '<button id="s1-close" class="w-9 h-9 rounded-xl bg-gray-100 flex items-center justify-center shrink-0"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#374151" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>' +
        '<div class="flex-1"><p class="font-semibold text-gray-900">Despacho de solicitud</p><p class="text-xs text-gray-400">Paso 1 de 2 — Encabezado</p></div>' +
        '<div class="flex gap-1.5"><div class="w-6 h-1.5 rounded-full" style="background:#1B4F8A"></div><div class="w-6 h-1.5 rounded-full bg-gray-200"></div></div>' +
      '</div>' +
      '<div class="flex-1 overflow-y-auto px-4 py-5 space-y-4">' +
        '<div>' +
          '<label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Usuario responsable *</label>' +
          '<div class="grid grid-cols-3 gap-2">' +
          USUARIOS_RESPONSABLES.map(u =>
            '<button data-resp="' + u + '" class="py-3.5 rounded-2xl border-2 text-sm font-bold transition-all ' + (hdr.responsable===u ? 'text-white border-transparent' : 'border-gray-200 text-gray-700 bg-white') + '" style="' + (hdr.responsable===u ? 'background:#1B4F8A' : '') + '">' + u + '</button>'
          ).join('') +
          '</div>' +
        '</div>' +
        '<div><label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Instalador responsable</label>' +
        '<select id="s1-instalador" class="w-full border-2 border-gray-200 rounded-2xl px-4 py-3 text-sm bg-white focus:outline-none focus:border-blue-400 font-medium text-gray-800">' +
          '<option value="">Seleccionar...</option>' +
          usuarios.map(u => '<option value="' + safeStr(u.displayName) + '" ' + (hdr.instalador===safeStr(u.displayName)?'selected':'') + '>' + safeStr(u.displayName) + '</option>').join('') +
        '</select></div>' +
        '<div class="grid grid-cols-2 gap-3">' +
          '<div><label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Empresa</label>' +
          '<select id="s1-contratista" class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400">' +
            EMPRESAS_CONTRATISTAS.map(e => '<option value="' + e + '" ' + (hdr.contratista===e?'selected':'') + '>' + e + '</option>').join('') +
            '<option value="" ' + (hdr.contratista===''?'selected':'') + '>Otra</option>' +
          '</select></div>' +
          '<div><label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">Placa</label>' +
          '<select id="s1-placa" class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400">' +
            '<option value="">Seleccionar...</option>' +
            PLACAS.map(p => '<option value="' + p + '" ' + (hdr.placaSel===p?'selected':'') + '>' + p + '</option>').join('') +
            '<option value="__otro__" ' + (hdr.placaSel==='__otro__'?'selected':'') + '>Otro</option>' +
          '</select>' +
          '<input id="s1-placa-otro" type="text" value="' + hdr.placaOtro + '" placeholder="Ej. P-123" class="' + (hdr.placaSel==='__otro__' ? '' : 'hidden') + ' w-full border-2 border-gray-200 rounded-2xl px-3 py-2.5 text-sm mt-2 font-mono focus:outline-none focus:border-blue-400"/>' +
          '</div>' +
        '</div>' +
        '<div class="grid grid-cols-2 gap-3">' +
          '<div><label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">F. solicitud</label><input id="s1-fecha-sol" type="date" value="' + hdr.fechaSol + '" class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400"/></div>' +
          '<div><label class="block text-xs font-semibold text-gray-500 uppercase tracking-wider mb-2">F. entrega</label><input id="s1-fecha-ent" type="date" value="' + hdr.fechaEnt + '" class="w-full border-2 border-gray-200 rounded-2xl px-3 py-3 text-sm bg-white focus:outline-none focus:border-blue-400"/></div>' +
        '</div>' +
      '</div>' +
      '<div class="px-4 py-4 border-t border-gray-100 shrink-0" style="padding-bottom:max(16px,env(safe-area-inset-bottom))">' +
        '<div id="s1-err" class="hidden text-sm text-red-600 text-center mb-3"></div>' +
        '<button id="s1-next" class="w-full font-bold rounded-2xl py-4 text-white flex items-center justify-center gap-2" style="background:#1B4F8A">Continuar — Materiales <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg></button>' +
      '</div>' +
    '</div>';

    ov.querySelectorAll('[data-resp]').forEach(b => { b.onclick = () => { hdr.responsable = b.dataset.resp; renderStep1(); }; });
    ov.querySelector('#s1-placa').addEventListener('change', e => {
      hdr.placaSel = e.target.value;
      const otro = ov.querySelector('#s1-placa-otro');
      if (e.target.value === '__otro__') { otro.classList.remove('hidden'); otro.focus(); }
      else { otro.classList.add('hidden'); hdr.placaOtro = ''; }
    });
    ov.querySelector('#s1-close').onclick = () => ov.remove();
    ov.querySelector('#s1-next').onclick = () => {
      hdr.instalador  = ov.querySelector('#s1-instalador').value;
      hdr.contratista = ov.querySelector('#s1-contratista').value;
      hdr.placaSel    = ov.querySelector('#s1-placa').value;
      hdr.placaOtro   = ov.querySelector('#s1-placa-otro')?.value.trim() || '';
      hdr.fechaSol    = ov.querySelector('#s1-fecha-sol').value;
      hdr.fechaEnt    = ov.querySelector('#s1-fecha-ent').value;
      if (!hdr.responsable) {
        ov.querySelector('#s1-err').textContent = 'Selecciona el usuario responsable.';
        ov.querySelector('#s1-err').classList.remove('hidden');
        return;
      }
      step = 2; renderStep2();
    };
  }

  function renderStep2() {
    const totalSel = sel.length;
    const placa = hdr.placaSel === '__otro__' ? hdr.placaOtro : hdr.placaSel;

    ov.innerHTML = '<div class="flex flex-col bg-gray-50" style="height:100dvh;max-height:100dvh">' +
      '<div class="flex items-center gap-3 px-4 py-3 border-b border-gray-100 bg-white shrink-0">' +
        '<button id="s2-back" class="w-9 h-9 rounded-xl bg-gray-100 flex items-center justify-center shrink-0"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#374151" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/></svg></button>' +
        '<div class="flex-1 min-w-0"><p class="font-semibold text-gray-900">Materiales</p><p class="text-xs text-gray-400 truncate">' + hdr.responsable + (hdr.instalador ? ' · ' + hdr.instalador : '') + (placa ? ' · ' + placa : '') + '</p></div>' +
        '<div class="flex gap-1.5"><div class="w-6 h-1.5 rounded-full bg-gray-200"></div><div class="w-6 h-1.5 rounded-full" style="background:#1B4F8A"></div></div>' +
      '</div>' +
      (totalSel > 0 ?
        '<div class="px-4 pt-3 pb-1 shrink-0 bg-white border-b border-gray-100">' +
          '<div class="flex items-center justify-between mb-2"><p class="text-xs font-semibold text-gray-500 uppercase tracking-wider">En el despacho</p><span class="text-xs font-bold px-2 py-0.5 rounded-full text-white" style="background:#1B4F8A">' + totalSel + '</span></div>' +
          '<div class="space-y-2 pb-1">' +
          sel.map(function(s, idx) {
            const restante  = s.stock - s.cantidad;
            const restColor = restante <= 0 ? '#C62828' : '#166534';
            return '<div class="flex items-center gap-3 bg-gray-50 rounded-2xl px-3 py-2.5 border border-gray-200">' +
              '<div class="flex-1 min-w-0"><p class="text-sm font-semibold text-gray-900 truncate leading-tight">' + tc(s.name) + '</p><p class="text-xs font-medium mt-0.5" style="color:' + restColor + '">Restante: ' + restante + ' ' + safeStr(s.unit,'') + '</p></div>' +
              '<div class="flex items-center gap-1.5 shrink-0">' +
                '<button data-dec="' + idx + '" class="w-8 h-8 rounded-xl border border-gray-300 bg-white text-lg font-bold text-gray-600 flex items-center justify-center">−</button>' +
                '<span class="w-8 text-center text-base font-bold text-gray-900">' + s.cantidad + '</span>' +
                '<button data-inc="' + idx + '" class="w-8 h-8 rounded-xl border text-lg font-bold flex items-center justify-center ' + (s.cantidad>=s.stock ? 'border-gray-200 text-gray-300 bg-gray-50' : 'border-blue-300 text-blue-600 bg-blue-50') + '" ' + (s.cantidad>=s.stock?'disabled':'') + '>+</button>' +
              '</div>' +
              '<button data-del="' + idx + '" class="w-7 h-7 rounded-xl bg-red-50 flex items-center justify-center shrink-0"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#EF4444" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>' +
            '</div>';
          }).join('') +
          '</div>' +
        '</div>'
      : '') +
      '<div class="flex-1 overflow-y-auto px-4 pt-3 pb-2">' +
        '<div class="relative mb-3"><svg class="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>' +
        '<input id="s2-buscar" type="text" value="' + busq + '" placeholder="Buscar material..." autocomplete="off" class="w-full bg-white border-2 border-gray-200 rounded-2xl pl-10 pr-4 py-3 text-sm focus:outline-none focus:border-blue-400"/></div>' +
        '<div id="s2-lista" class="space-y-1.5 pb-2"></div>' +
      '</div>' +
      '<div id="s2-err" class="hidden mx-4 mb-1 text-sm text-red-600 text-center shrink-0"></div>' +
      '<div class="px-4 py-3 bg-white border-t border-gray-200 shrink-0" style="padding-bottom:max(16px,env(safe-area-inset-bottom))">' +
        '<button id="s2-submit" class="w-full font-bold rounded-2xl py-4 text-sm text-white transition-all" style="background:' + (totalSel>0 ? '#1B4F8A' : '#D1D5DB') + '" ' + (totalSel===0?'disabled':'') + '>' +
          (totalSel>0 ? 'Registrar salida · ' + totalSel + ' material' + (totalSel!==1?'es':'') : 'Agrega materiales para continuar') +
        '</button>' +
      '</div>' +
    '</div>';

    ov.querySelector('#s2-back').onclick = () => { step=1; renderStep1(); };
    ov.querySelectorAll('[data-dec]').forEach(b => { b.onclick = () => { if(sel[+b.dataset.dec].cantidad>1){sel[+b.dataset.dec].cantidad--;renderStep2();}};});
    ov.querySelectorAll('[data-inc]').forEach(b => { b.onclick = () => { if(sel[+b.dataset.inc].cantidad<sel[+b.dataset.inc].stock){sel[+b.dataset.inc].cantidad++;renderStep2();}};});
    ov.querySelectorAll('[data-del]').forEach(b => { b.onclick = () => { sel.splice(+b.dataset.del,1);renderStep2();};});
    ov.querySelector('#s2-buscar').addEventListener('input', e => { busq=e.target.value; renderLista2(); });
    ov.querySelector('#s2-submit')?.addEventListener('click', handleSubmit);
    renderLista2();
  }

  function renderLista2() {
    const el = ov.querySelector('#s2-lista');
    if (!el) return;
    const selIds = new Set(sel.map(s=>s.itemId));
    const q = busq.trim().toLowerCase();
    const lista = q ? items.filter(i => safeStr(i.name).toLowerCase().includes(q) || safeStr(i.sapCode,'').includes(q)) : items;
    if (!lista.length) { el.innerHTML = '<p class="text-sm text-gray-400 text-center py-8">Sin resultados</p>'; return; }
    el.innerHTML = lista.map(item => {
      const agregado = selIds.has(item.id);
      return '<button data-item="' + item.id + '" ' + (agregado?'disabled':'') + ' class="w-full flex items-center justify-between gap-3 text-left px-4 py-3.5 rounded-2xl border-2 ' + (agregado ? 'border-green-200 bg-green-50 cursor-not-allowed' : 'border-gray-200 bg-white active:border-blue-400 active:bg-blue-50') + '">' +
        '<p class="text-sm font-semibold truncate ' + (agregado?'text-green-700':'text-gray-900') + '">' + tc(item.name) + '</p>' +
        (agregado ? '<span class="text-xs font-bold text-green-600 shrink-0">✓</span>' : '<span class="text-sm font-bold shrink-0" style="color:#374151">' + item.stock + ' <span class="text-xs font-normal text-gray-400">' + safeStr(item.unit,'') + '</span></span>') +
      '</button>';
    }).join('');
    el.querySelectorAll('[data-item]').forEach(btn => {
      btn.addEventListener('click', () => {
        const item = items.find(i=>i.id===btn.dataset.item);
        if (!item || sel.some(s=>s.itemId===item.id)) return;
        showModalCant(item);
      });
    });
  }

  function showModalCant(item) {
    const m = document.createElement('div');
    m.className = 'fixed inset-0 bg-black/50 flex items-end z-50';
    m.innerHTML = '<div class="bg-white w-full rounded-t-3xl px-5 pt-5 pb-8 space-y-4" style="padding-bottom:max(32px,env(safe-area-inset-bottom))">' +
      '<div class="w-10 h-1 bg-gray-200 rounded-full mx-auto"></div>' +
      '<div><p class="font-bold text-gray-900 text-lg">' + tc(item.name) + '</p><p class="text-sm text-gray-400">' + item.stock + ' ' + safeStr(item.unit,'') + ' disponibles</p></div>' +
      '<div class="flex items-center justify-between gap-4">' +
        '<button id="mc2-dec" class="w-16 h-16 rounded-2xl border-2 border-gray-200 bg-gray-50 text-3xl font-bold text-gray-400 flex items-center justify-center">−</button>' +
        '<div class="flex-1 text-center"><input id="mc2-cant" type="number" min="1" max="' + item.stock + '" value="1" class="w-full text-center text-5xl font-black text-gray-900 bg-transparent border-none focus:outline-none"/><p class="text-sm text-gray-400">' + safeStr(item.unit,'') + '</p></div>' +
        '<button id="mc2-inc" class="w-16 h-16 rounded-2xl border-2 border-blue-300 bg-blue-50 text-3xl font-bold text-blue-600 flex items-center justify-center">+</button>' +
      '</div>' +
      '<div id="mc2-err" class="hidden text-sm text-red-500 text-center"></div>' +
      '<button id="mc2-add" class="w-full text-white font-bold rounded-2xl py-4" style="background:#1B4F8A">Agregar al despacho</button>' +
    '</div>';
    document.body.appendChild(m);
    const cantEl = m.querySelector('#mc2-cant');
    setTimeout(() => { cantEl.focus(); cantEl.select(); }, 80);
    m.addEventListener('click', e => { if(e.target===m) m.remove(); });
    m.querySelector('#mc2-dec').onclick = () => { const v=safeNum(cantEl.value); if(v>1) cantEl.value=v-1; };
    m.querySelector('#mc2-inc').onclick = () => { const v=safeNum(cantEl.value); if(v<item.stock) cantEl.value=v+1; };
    const doAdd = () => {
      const cant=safeNum(cantEl.value);
      const err=m.querySelector('#mc2-err');
      if(cant<=0){err.textContent='Ingresa una cantidad mayor a 0.';err.classList.remove('hidden');return;}
      if(cant>item.stock){err.textContent='Máximo: '+item.stock;err.classList.remove('hidden');return;}
      sel.push({itemId:item.id,name:item.name,unit:item.unit,sapCode:item.sapCode,axCode:item.axCode,stock:item.stock,cantidad:cant,requiereSerial:item.requiereSerial,modoSerial:'individual',seriales:[],serialInicio:'',serialFin:''});
      addRec(item.id); m.remove(); renderStep2();
    };
    m.querySelector('#mc2-add').addEventListener('click', doAdd);
    cantEl.addEventListener('keydown', e => { if(e.key==='Enter') doAdd(); });
  }

  async function handleSubmit() {
    const errEl = ov.querySelector('#s2-err');
    const btn   = ov.querySelector('#s2-submit');
    if (!sel.length) return;
    errEl?.classList.add('hidden');
    btn.disabled = true; btn.textContent = 'Registrando...';
    const placa = hdr.placaSel==='__otro__' ? hdr.placaOtro : hdr.placaSel;
    try {
      const ref = await addDoc(collection(db, 'kardex/movimientos/salidas'), {
        usuarioResponsableUid: hdr.responsable,
        usuarioResponsable:    hdr.responsable,
        empresaContratista:    hdr.contratista,
        instaladorResponsable: hdr.instalador,
        placaVehiculo:         placa,
        fechaSolicitud: hdr.fechaSol,
        fechaEntrega:   hdr.fechaEnt,
        entregadoPor:    session.displayName,
        entregadoPorUid: session.uid,
        solicitudId:     solicitudId,
        items: sel.map(s => ({
          itemId:s.itemId, sapCode:s.sapCode, axCode:s.axCode,
          nombre:s.name, unit:s.unit, cantidad:s.cantidad,
          requiereSerial:s.requiereSerial,
          modoSerial:s.requiereSerial?s.modoSerial:null,
          seriales:s.requiereSerial&&s.modoSerial==='individual'?s.seriales:[],
          serialInicio:s.requiereSerial&&s.modoSerial==='rango'?s.serialInicio:'',
          serialFin:s.requiereSerial&&s.modoSerial==='rango'?s.serialFin:'',
        })),
        fecha: serverTimestamp(),
      });
      for (const s of sel) {
        await updateDoc(doc(db,'kardex/inventario/items',s.itemId),{stock:increment(-s.cantidad)});
      }
      // Marcar solicitud como despachada si viene de una
      if (solicitudId) {
        await updateDoc(doc(db,'solicitudes_material',solicitudId),{
          estado: 'aprobado',
          salidaId: ref.id,
          fechaDespacho: serverTimestamp(),
        });
      }
      ov.remove();
      showToast('Salida registrada correctamente.','success');
      showMemo({
        id:ref.id, usuarioResponsable:hdr.responsable,
        empresaContratista:hdr.contratista, instaladorResponsable:hdr.instalador,
        entregadoPor:session.displayName, placaVehiculo:placa,
        fechaSolicitud:hdr.fechaSol, fechaEntrega:hdr.fechaEnt, items:sel,
      });
      await showDashboard(db, session);
    } catch(e) {
      if(errEl){errEl.textContent='Error al registrar. Intenta de nuevo.';errEl.classList.remove('hidden');}
      btn.disabled=false;
      btn.textContent='Registrar salida · '+sel.length+' material'+(sel.length!==1?'es':'');
      console.error(e);
    }
  }

  if (step === 2) renderStep2(); else renderStep1();
}
