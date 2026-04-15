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
// INIT
// ─────────────────────────────────────────
export async function initKardex(session) {
  const container = document.getElementById('kardex-root');
  if (!container) return;
  const db = window.__firebase.db;
  renderShell(container, session);
  await showDashboard(db, session);
  bindNav(db, session);
}

// ─────────────────────────────────────────
// SHELL
// ─────────────────────────────────────────
function renderShell(container, session) {
  const canEdit = ['admin','coordinadora'].includes(session.role);
  container.innerHTML = `
    <div class="space-y-4">
      <div class="flex items-center justify-between">
        <div>
          <h1 class="text-xl font-semibold text-gray-900">Kardex</h1>
          <p class="text-sm text-gray-500 mt-0.5">Control de materiales</p>
        </div>
        ${canEdit ? `
        <button id="btn-nueva-salida"
          class="inline-flex items-center gap-2 text-white text-sm font-medium px-4 py-2.5 rounded-lg"
          style="background-color:#1B4F8A">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
          </svg>
          Nueva salida
        </button>` : ''}
      </div>
      <div class="flex gap-1 bg-gray-100 rounded-xl p-1">
        <button data-tab="dashboard"  class="ktab flex-1 py-2 text-sm font-medium rounded-lg transition-colors">Inicio</button>
        <button data-tab="inventario" class="ktab flex-1 py-2 text-sm font-medium rounded-lg transition-colors">Inventario</button>
        <button data-tab="historial"  class="ktab flex-1 py-2 text-sm font-medium rounded-lg transition-colors">Historial</button>
      </div>
      <div id="kardex-content"></div>
    </div>`;
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
      if (t === 'dashboard')  await showDashboard(db, session);
      if (t === 'inventario') await showInventario(db, session);
      if (t === 'historial')  await showHistorial(db, session);
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
        if (t === 'dashboard')  await showDashboard(db, session);
        if (t === 'inventario') await showInventario(db, session);
        if (t === 'historial')  await showHistorial(db, session);
      });
    });
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function statCard(val, label, color) {
  return `<div class="bg-white rounded-xl border border-gray-200 p-4 text-center">
    <p class="text-3xl font-bold" style="color:${color}">${val}</p>
    <p class="text-xs text-gray-500 mt-1">${label}</p>
  </div>`;
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
  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Registrar entrada</h2>
      <button id="fe-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div class="px-5 py-4 space-y-4">
      <div class="bg-gray-50 rounded-xl p-3">
        <p class="text-sm font-medium text-gray-900">${safeStr(item?.name)}</p>
        <p class="text-xs text-gray-400 font-mono mt-0.5">
          ${item?.sapCode ? `SAP ${item.sapCode}` : ''}${item?.sapCode && item?.axCode ? ' · ' : ''}${item?.axCode ? `AX ${item.axCode}` : ''}
        </p>
        <p class="text-xs text-gray-500 mt-1">Stock actual: <strong>${safeNum(item?.stock)} ${safeStr(item?.unit,'')}</strong></p>
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Cantidad que ingresa</label>
        <input id="fe-cant" type="number" min="1" placeholder="0"
          class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 text-center text-lg font-semibold"/>
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Motivo</label>
        <input id="fe-motivo" type="text" placeholder="Ej. Compra, Reposición"
          class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
      </div>
      <div id="fe-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fe-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cancelar</button>
      <button id="fe-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#2E7D32">Registrar entrada</button>
    </div>`);

  ov.querySelector('#fe-close').onclick = ov.querySelector('#fe-cancel').onclick = () => ov.remove();
  ov.querySelector('#fe-submit').onclick = async () => {
    const cant   = safeNum(ov.querySelector('#fe-cant').value);
    const motivo = ov.querySelector('#fe-motivo').value.trim();
    const errEl  = ov.querySelector('#fe-err');
    const btn    = ov.querySelector('#fe-submit');
    errEl.classList.add('hidden');
    if (cant <= 0) { errEl.textContent = 'Cantidad inválida.'; errEl.classList.remove('hidden'); return; }
    btn.disabled = true; btn.textContent = 'Registrando...';
    try {
      const stockAntes = safeNum(item?.stock);
      await updateDoc(doc(db, 'kardex/inventario/items', item.id), { stock: increment(cant) });
      await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
        itemId: item.id, itemNombre: safeStr(item?.name), tipo: 'entrada', cantidad: cant, motivo,
        stockAntes, stockDespues: stockAntes + cant,
        fecha: serverTimestamp(), registradoPor: session.uid, registradoPorNombre: session.displayName,
      });
      ov.remove();
      showToast(`Entrada de ${cant} ${safeStr(item?.unit,'')} registrada.`, 'success');
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

  // ── Modal cantidad ──
  function showCantidadModal(item) {
    const modal = document.createElement('div');
    modal.className = 'fixed inset-0 bg-black/50 flex items-end z-50';
    modal.innerHTML = `
      <div class="bg-white w-full rounded-t-3xl px-5 pt-5 space-y-5"
        style="padding-bottom:max(32px,env(safe-area-inset-bottom))">
        <div class="w-10 h-1 bg-gray-200 rounded-full mx-auto mb-1"></div>
        <div>
          <p class="font-bold text-gray-900 text-lg leading-tight">${tc(item.name)}</p>
          <p class="text-sm text-gray-400 mt-0.5">${item.stock} ${safeStr(item.unit,'')} disponibles</p>
        </div>
        <div class="flex items-center justify-between gap-4">
          <button id="mc-dec"
            class="w-16 h-16 rounded-2xl border-2 border-gray-200 bg-gray-50 text-3xl font-bold text-gray-400 flex items-center justify-center active:bg-gray-200">−</button>
          <div class="flex-1 text-center">
            <input id="mc-cant" type="number" min="1" max="${item.stock}" value="1"
              class="w-full text-center text-5xl font-black text-gray-900 bg-transparent border-none focus:outline-none leading-none py-2"/>
            <p class="text-sm text-gray-400 -mt-1">${safeStr(item.unit,'')}</p>
          </div>
          <button id="mc-inc"
            class="w-16 h-16 rounded-2xl border-2 border-blue-300 bg-blue-50 text-3xl font-bold text-blue-600 flex items-center justify-center active:bg-blue-100">+</button>
        </div>
        <div id="mc-err" class="hidden text-sm text-red-500 text-center"></div>
        <button id="mc-add"
          class="w-full text-white font-bold rounded-2xl py-4 text-base active:opacity-90" style="background:#1B4F8A">
          Agregar al despacho
        </button>
      </div>`;

    document.body.appendChild(modal);
    const cantEl = modal.querySelector('#mc-cant');
    setTimeout(() => { cantEl.focus(); cantEl.select(); }, 80);

    modal.addEventListener('click', e => { if (e.target===modal) modal.remove(); });
    modal.querySelector('#mc-dec').onclick = () => { const v=safeNum(cantEl.value); if(v>1) cantEl.value=v-1; };
    modal.querySelector('#mc-inc').onclick = () => { const v=safeNum(cantEl.value); if(v<item.stock) cantEl.value=v+1; };

    const doAdd = () => {
      const cant  = safeNum(cantEl.value);
      const errEl = modal.querySelector('#mc-err');
      if (cant<=0)        { errEl.textContent='Ingresa una cantidad mayor a 0.'; errEl.classList.remove('hidden'); return; }
      if (cant>item.stock){ errEl.textContent=`Máximo: ${item.stock} ${safeStr(item.unit,'')}`; errEl.classList.remove('hidden'); return; }
      sel.push({
        itemId:item.id, name:item.name, unit:item.unit,
        sapCode:item.sapCode, axCode:item.axCode, stock:item.stock,
        cantidad:cant, requiereSerial:item.requiereSerial,
        modoSerial:'individual', seriales:[], serialInicio:'', serialFin:'',
      });
      addRec(item.id);
      modal.remove();
      renderStep2();
    };

    modal.querySelector('#mc-add').addEventListener('click', doAdd);
    cantEl.addEventListener('keydown', e => { if(e.key==='Enter') doAdd(); });
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
async function showHistorial(db, session) {
  setTab('historial');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  try {
    const q     = query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'));
    const snap  = await getDocs(q);
    const salidas = snap.docs.map(d => ({ id: d.id, ...d.data() }));

    content.innerHTML = `
      <div class="bg-white rounded-xl border border-gray-200 overflow-hidden">
        <div class="px-4 py-3 border-b border-gray-100">
          <p class="font-semibold text-sm text-gray-900">Historial de salidas</p>
          <p class="text-xs text-gray-400 mt-0.5">${salidas.length} registro(s)</p>
        </div>
        ${salidas.length === 0
          ? '<div class="py-12 text-center text-sm text-gray-400">Sin movimientos registrados</div>'
          : '<div class="divide-y divide-gray-100">' + salidas.map(s => `
            <div class="px-4 py-3">
              <div class="flex items-start justify-between gap-2">
                <div class="min-w-0">
                  <p class="text-sm font-semibold text-gray-900">${safeStr(s.usuarioResponsable || s.tecnicoNombre)}</p>
                  <p class="text-xs text-gray-400">${fmtDate(s.fecha)} · ${safeStr(s.entregadoPor || s.registradoPorNombre)}</p>
                  ${s.empresaContratista ? `<p class="text-xs text-gray-400">${s.empresaContratista}</p>` : ''}
                </div>
                <button data-smemo="${s.id}" class="text-xs font-medium px-2 py-1 rounded-lg shrink-0" style="color:#1B4F8A;background:#EFF6FF">Memo</button>
              </div>
              <div class="flex flex-wrap gap-1 mt-2">
                ${(s.items||[]).map(i =>
                  `<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">
                    ${safeNum(i.cantidad)} ${safeStr(i.unit||i.unidad,'')} ${safeStr(i.nombre||i.name)}
                  </span>`).join('')}
              </div>
            </div>`).join('') + '</div>'
        }
      </div>`;

    const salidaMap = {};
    salidas.forEach(s => { salidaMap[s.id] = s; });
    content.querySelectorAll('[data-smemo]').forEach(b => {
      b.onclick = () => {
        const s = salidaMap[b.dataset.smemo];
        if (s) showMemo({ ...s, fecha: s.fecha?.toDate?.() || new Date() });
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
      <button id="fm-print"  class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">🖨️ Generar PDF</button>
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
// IMPRIMIR DESPACHO OFICIAL
// HTML fiel al documento Word original.
// Medidas exactas extraídas del .docx:
//   Página carta: 215.9×279.4mm
//   Márgenes: top=8.8 bottom=4.9 left=12.7 right=6.3
//   Tabla encabezado: 2 cols de 50% c/u
//   Tabla materiales: 10.6% | 7.0% | 72.8% | 9.6%
//   Fuente header: 8pt bold, datos: 7.5pt normal
// ─────────────────────────────────────────

// 65 filas de la tabla 1 del Word, en orden exacto
const DOC_ROWS = [
  {sap:"RESERVA",ax:"STOCK",desc:"DESCRIPICIÓN",header:true},
  {sap:"USO HABITUAL",ax:"",desc:"",header:true},
  {sap:"221477",ax:"50203",desc:"ALAMBRE COBRE THHN 8 AWG 600 V FORRO PLASTICO",header:false},
  {sap:"213719",ax:"50806",desc:"CABLE DUPLEX AL #6 ACSR SETTER",header:false},
  {sap:"328541",ax:"50807",desc:"CABLE TRIPLEX AL. #6 ACSR PALUDINA",header:false},
  {sap:"352453",ax:"250201",desc:"CONECTOR DE COMPRESIÓN YPC2A8U",header:false},
  {sap:"352460",ax:"250202",desc:"CONECTOR DE COMPRESIÓN YPC26R8U",header:false},
  {sap:"352461",ax:"250203",desc:"CONECTOR DE COMPRESIÓN YP2U3",header:false},
  {sap:"352462",ax:"250204",desc:"CONECTOR DE COMPRESIÓN YP26AU2",header:false},
  {sap:"353112",ax:"400910",desc:"ANCLA PLASTICA 1 1/2 X 7 (FTN1-120)",header:false},
  {sap:"354045",ax:"400919",desc:"TORNILLO CABEZA PLANA DE 11/2 PLG X 7MM (Gruesa de 144 unidades)",header:false},
  {sap:"354549",ax:"400931",desc:"SELLO ACRILICO VERDE (SERV. NVOS., MTTO.) (CABLE 30 CM) (FTMED-30)",header:false},
  {sap:"200129",ax:"700101",desc:"MEDIDOR BIFILAR DOMICILIAR BASE A, 100 A (ETE1-330)",header:false},
  {sap:"355518",ax:"700102",desc:"MEDIDOR TRIFILAR DOMICILIAR BASE A, 100 A (ETE1-330)",header:false},
  {sap:"338362",ax:"700326",desc:"MEDIDOR FORMA 2(S) T/ESPIGA, CLASE 100, TRIFI. 240 V, 15/100",header:false},
  {sap:"219359",ax:"750109",desc:"CINTA AISLANTE SUPER 3M #33",header:false},
  {sap:"MATERIAL PARA CL200",ax:"",desc:"",header:true},
  {sap:"328560",ax:"50205",desc:"CABLE COBRE THHN # 2 AWG 600 V FORRO PLASTICO (ETM3-310)",header:false},
  {sap:"243940",ax:"50209",desc:"CABLE COBRE THHN # 1/0 AWG 19 HILOS (ETM3-310)",header:false},
  {sap:"337775",ax:"250101",desc:"CONECTOR MECANICO PERNO PARTIDO KSU-23",header:false},
  {sap:"337776",ax:"250102",desc:"CONECTOR MECANICO PERNO PARTIDO KSU-26",header:false},
  {sap:"337777",ax:"250103",desc:"CONECTOR MECANICO PERNO PARTIDO KSU-29",header:false},
  {sap:"210525",ax:"400833",desc:"TUBO EMT 2 PLG ALUMINIO UL (FTN1-700)",header:false},
  {sap:"212720",ax:"400834",desc:"CUERPO TERMINAL PARA EMT DE 2 PLG CON ABRAZADERA",header:false},
  {sap:"214221",ax:"400836",desc:"CONECTOR EMT DE 2 PLG CON TORNILLO",header:false},
  {sap:"301560",ax:"400837",desc:"GRAPA CONDUIT EMT DE 2 PLG",header:false},
  {sap:"245979",ax:"400838",desc:"TUBO EMT DE 2 1/2 PLG (FTN1-700)",header:false},
  {sap:"355070",ax:"400839",desc:"CUERPO TERMINAL PARA EMT DE 2 1/2 PLG CON ABRAZADERA",header:false},
  {sap:"244569",ax:"400840",desc:"CONECTOR EMT DE 2 1/2 PLG CON TORNILLO",header:false},
  {sap:"353992",ax:"400841",desc:"GRAPA CONDUIT EMT DE 2 1/2 PLG",header:false},
  {sap:"211373",ax:"400903",desc:"CINTA BAND IT 3/4",header:false},
  {sap:"211375",ax:"400904",desc:"HEBILLA PARA CINTA BAND IT 3/4",header:false},
  {sap:"353121",ax:"700406",desc:"TAPADERA DE POLICARBONATO PARA BASES SOCKET TIPO ESPIGA",header:false},
  {sap:"355064",ax:"700332",desc:"MEDIDOR FORMA 16s, CLASE 200, 120-277V. 8 CANALES DE MEM. 200 AMP. C/BASE 7 TERMI. (ETE-16s)",header:false},
  {sap:"338357",ax:"700333",desc:"MEDIDOR FORMA 12s CLASE 200, 120-277V, TRIFILAR, 60Hz, 8 Canales de memoria, C/Base 5 Term (ETE-12s)",header:false},
  {sap:"353099",ax:"401102",desc:"BASES SOCKETS DE 5 TERMINALES, 200 AMP",header:false},
  {sap:"353110",ax:"700404",desc:"BASES SOCKETS 100 AMP. P/MED ESPIGA F 2(S)",header:false},
  {sap:"353088",ax:"401101",desc:"BASES SOCKETS DE 7 TERMINALES MARCA THOMAS & BETTS, 200 AMPS",header:false},
  {sap:"ACOMETIDA ESPECIAL",ax:"",desc:"",header:true},
  {sap:"200468",ax:"50401",desc:"CABLE DE ALUMINIO #2 AWG F.P. 7 HILOS (ETM3-330)",header:false},
  {sap:"200472",ax:"50809",desc:"CABLE DE ALUMINIO ACSR DESNUDO #2 AWG 7 HILOS SPARROW (ETM3-350)",header:false},
  {sap:"200469",ax:"50402",desc:"CABLE DE ALUMINIO #1/0 AWG F.P. 7 HILOS (ETM3-330)",header:false},
  {sap:"200473",ax:"50810",desc:"CABLE DE ALUMINIO ACSR DESNUDO #1/0 AWG 7 HILOS, (ETM3-350)",header:false},
  {sap:"213410",ax:"50505",desc:"CABLE TRIPLEX ACSR 1/0 AWG NERITINA (ETM3-330)",header:false},
  {sap:"214726",ax:"150202",desc:"REMATE PREFORMADO ACSR 2 AWG (ETM1-240)",header:false},
  {sap:"214727",ax:"150205",desc:"REMATE PREFORMADO ACSR 1/0 AWG (ETM1-240)",header:false},
  {sap:"352463",ax:"250205",desc:"CONECTOR MECANICO COMPRESIÓN YP25U25",header:false},
  {sap:"SUBTERRÁNEO",ax:"",desc:"",header:true},
  {sap:"219527",ax:"250418",desc:"TERMINAL DE OJO CABLE 4 AWG, DIAMETRO 3/8 PLG (ETM1-460)",header:false},
  {sap:"221062",ax:"250420",desc:"TERMINAL DE OJO 1/0 DIAMETRO 3/8 PLG (FTN1-320)",header:false},
  {sap:"200367",ax:"50220",desc:"CABLE COBRE XHHW # 4 AWG 19 HILOS (ETM3-310)",header:false},
  {sap:"282485",ax:"50219",desc:"CABLE COBRE XHHW # 6 AWG 7 HILOS (ETM3-310)",header:false},
  {sap:"350560",ax:"50231",desc:"CABLE COBRE RHHW #1/0 AWG",header:false},
  {sap:"212896",ax:"250419",desc:"TERMINAL DE OJO NO. 2 DIAMETRO 3/8 PLG",header:false},
  {sap:"350564",ax:"50234",desc:"CABLE COBRE RHHW # 2 AWG 19 HILOS",header:false},
  {sap:"PATRON ANTIHURTO Y TELEGESTIÓN",ax:"",desc:"",header:true},
  {sap:"221472",ax:"50710",desc:"CABLE CONCENTRICO TELESCOPICO CCA BIFILAR 6 AWG PARA ACOMETIDA AEREA (ETM3-470)",header:false},
  {sap:"200413",ax:"50721",desc:"CABLE CONCENTRICO TELESCOPICO CCA TRIFILAR 6 AWG PARA ACOMETIDA AEREA (ETM3-470)",header:false},
  {sap:"211829",ax:"400707",desc:"CAJA TRANSPARENTE DE POLICARBONATO PARA MEDIDOR TOTALIZADOR",header:false},
  {sap:"213340",ax:"150419",desc:"PINZAS DE RETENCION PARA CABLE CONCENTRICO TELESCOPICO #1/0 AWG (FTN1-760)",header:false},
  {sap:"222315",ax:"150420",desc:"PINZAS DE RETENCION PARA CABLE CONCENTRICO TELESCOPICO DE ACOMETIDA BT #6 (FTN1-770)",header:false},
  {sap:"353730",ax:"750116",desc:"CINTA DE BLINDAJE 3M, CAT. ARMORCAST 4560-10",header:false},
  {sap:"338363",ax:"700340",desc:"MEDIDOR TELEGESTIONADO BASE A - 240V (FTMED-32)",header:false},
  {sap:"338361",ax:"700339",desc:"MEDIDOR RESIDENCIAL PREPAGO CLASE 100, 120V CON TELEGESTION",header:false},
  {sap:"338360",ax:"700338",desc:"MEDIDOR RESIDENCIAL POSTPAGO CLASE 100, 120-240V CON TELEGESTION",header:false},
];

function imprimirDespacho(memo) {
  const cantMap = {};
  for (const item of memo.MATERIALES) {
    cantMap[String(item.RESERVA).trim()] = item.CANTIDAD;
  }

  const filas = DOC_ROWS.map(row => {
    if (row.header) {
      if (row.sap === 'RESERVA') {
        return `<tr><td class="th">RESERVA</td><td class="th">STOCK</td><td class="th cant">CANTIDAD</td><td class="th desc">DESCRIPICIÓN</td></tr>`;
      }
      return `<tr><td class="sec" colspan="4">${row.sap}</td></tr>`;
    }
    const cant = cantMap[row.sap] || '';
    return `<tr><td class="c0">${row.sap}</td><td class="c1">${row.ax}</td><td class="c2 cant">${cant}</td><td class="c3">${row.desc}</td></tr>`;
  }).join('');

  // Seriales
  const seriales = memo.SERIALES || [];
  let serialHtml = '';
  if (seriales.length > 0) {
    const sRows = seriales.map(i => {
      if (i.modoSerial === 'individual' && (i.seriales||[]).length > 0) {
        return i.seriales.map((ser, idx) => `<tr>
          <td class="c0">${idx===0 ? safeStr(i.axCode,'') : ''}</td>
          <td class="c0">${idx===0 ? safeStr(i.sapCode,'') : ''}</td>
          <td>${idx===0 ? safeStr(i.nombre||i.name,'').toUpperCase() : ''}</td>
          <td class="cant">${idx===0 ? safeNum(i.cantidad) : ''}</td>
          <td class="cant">${idx+1}</td>
          <td class="c0">${ser}</td>
          <td></td>
        </tr>`).join('');
      }
      const ini = i.modoSerial==='rango' ? (i.serialInicio||'') : (i.seriales||[])[0]||'';
      const fin = i.modoSerial==='rango' ? (i.serialFin||'')   : (i.seriales||[]).slice(-1)[0]||'';
      return `<tr>
        <td class="c0">${safeStr(i.axCode,'')}</td>
        <td class="c0">${safeStr(i.sapCode,'')}</td>
        <td>${safeStr(i.nombre||i.name,'').toUpperCase()}</td>
        <td class="cant">${safeNum(i.cantidad)}</td>
        <td class="cant">1</td>
        <td class="c0">${ini}</td>
        <td class="c0">${fin}</td>
      </tr>`;
    }).join('');
    serialHtml = `
      <p class="serial-title">Serial de medidores / sellos entregados</p>
      <table>
        <tr><td class="th">STOCK</td><td class="th">RESERVA</td><td class="th desc">DESCRIPCIÓN</td><td class="th cant">CANTIDAD</td><td class="th cant">Cantidad</td><td class="th">Inicio</td><td class="th">Fin</td></tr>
        ${sRows}
      </table>`;
  }

  const v = memo;
  const html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<title>Despacho / Carga de Materiales</title>
<style>
/* Página carta con los márgenes exactos del Word */
@page { size: 215.9mm 279.4mm; margin: 8.8mm 6.3mm 4.9mm 12.7mm; }
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family: Arial, sans-serif; font-size: 7.5pt; color: #000; width: 196.9mm; }

/* ── Tabla encabezado institucional ──
   Replica Tabla 0 del Word: 2 cols de 50%, bordes negros, sin fondo */
.t0 { width: 100%; border-collapse: collapse; border: 1px solid #000; margin-bottom: 1mm; }
.t0 td { border: 1px solid #000; padding: 0.8mm 1mm; vertical-align: top; font-size: 7.5pt; }
.empresa { font-size: 8pt; font-weight: bold; text-align: left; }
.otc     { font-size: 7.5pt; }
.label   { font-size: 6.5pt; color: #000; display: block; }
.valor   { font-size: 7.5pt; display: block; min-height: 3.5mm; border-bottom: 0.5pt solid #555; padding-bottom: 0.3mm; margin-top: 0.2mm; }

/* ── Tabla materiales ──
   Replica Tabla 1 del Word: 4 cols, proporciones exactas */
.t1 { width: 100%; border-collapse: collapse; }
.t1 td { border: 1px solid #000; padding: 0.3mm 0.8mm; font-size: 7.5pt; vertical-align: middle; }
/* Anchos exactos del Word */
.c0   { width: 10.6%; text-align: center; }
.c1   { width: 7.0%;  text-align: center; }
.c2   { width: 9.6%;  text-align: center; }
.c3   { width: 72.8%; }
/* Alias para tabla seriales */
.cant { text-align: center; }
.desc { }
/* Encabezado columnas */
.th  { background: #d9d9d9; font-weight: bold; font-size: 8pt; text-align: center; border: 1px solid #000; padding: 0.5mm 0.8mm; }
/* Headers de sección */
.sec { background: #d9d9d9; font-weight: bold; font-size: 7.5pt; border: 1px solid #000; padding: 0.3mm 0.8mm; }

/* Firmas */
.firmas { display: flex; margin-top: 12mm; gap: 10mm; }
.firma  { flex:1; text-align: center; }
.firma .linea { border-top: 0.75pt solid #000; padding-top: 1mm; font-size: 7pt; font-weight: bold; }

.serial-title { font-size: 7.5pt; font-weight: bold; text-transform: uppercase; margin: 3mm 0 1mm; }

@media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
</style></head><body>

<!-- Tabla 0: Encabezado institucional -->
<!-- Fila 0-2 col izquierda: empresa, OTC, DESPACHO — unidas con rowspan -->
<!-- Fila 0 col derecha: vacía -->
<!-- Filas 3-7: labels y valores en ambas columnas -->
<table class="t0">
  <colgroup><col style="width:50%"><col style="width:50%"></colgroup>
  <tr>
    <td rowspan="3" style="vertical-align:middle">
      <span class="empresa">DISTRUIBUIDORA DE ELECTRICIDAD DELSUR S.A. DE C.V.</span><br>
      <span class="otc">OTC - GERENCIA COMERCIAL</span><br>
      <span class="otc">DESPACHO/ CARGA DE MATERIALES</span>
    </td>
    <td><span class="label">USUARIO RESPONSABLE:</span><span class="valor">${v.USUARIO_RESPONSABLE}</span></td>
  </tr>
  <tr>
    <td>
      <table style="width:100%;border-collapse:collapse"><tr>
        <td style="width:50%;padding-right:1mm;border:none">
          <span class="label">EMPRESA CONTRATISTA:</span><span class="valor">${v.EMPRESA_CONTRATISTA}</span>
        </td>
        <td style="width:50%;padding-left:1mm;border:none;border-left:0.5pt solid #999">
          <span class="label">INSTALADOR RESPONSABLE:</span><span class="valor">${v.INSTALADOR_RESPONSABLE}</span>
        </td>
      </tr></table>
    </td>
  </tr>
  <tr>
    <td>
      <table style="width:100%;border-collapse:collapse"><tr>
        <td style="width:50%;padding-right:1mm;border:none">
          <span class="label">ENTREGADO POR:</span><span class="valor">${v.ENTREGADO_POR}</span>
        </td>
        <td style="width:50%;padding-left:1mm;border:none;border-left:0.5pt solid #999">
          <span class="label">FIRMA DE RECIBIDO:</span><span class="valor">&nbsp;</span>
        </td>
      </tr></table>
    </td>
  </tr>
  <tr>
    <td><span class="label">FIRMA DE ENTREGADO:</span><span class="valor">&nbsp;</span></td>
    <td><span class="label">PLACA DE VEHICULO:</span><span class="valor">${v.PLACA_VEHICULO}</span></td>
  </tr>
  <tr>
    <td><span class="label">FECHA DE SOLICITUD</span><span class="valor">${v.FECHA_SOLICITUD}</span></td>
    <td><span class="label">FECHA ENTREGA DE MATERIAL:</span><span class="valor">${v.FECHA_ENTREGA}</span></td>
  </tr>
</table>

<!-- Tabla 1: Materiales con todos los 59 materiales y secciones -->
<table class="t1">
  <colgroup>
    <col class="c0"><col class="c1"><col class="c2"><col class="c3">
  </colgroup>
  ${filas}
</table>

${serialHtml}

<!-- Firmas -->
<div class="firmas">
  <div class="firma"><div class="linea">FIRMA DE ENTREGADO</div></div>
  <div class="firma"><div class="linea">FIRMA DE RECIBIDO</div></div>
</div>

<script>window.onload=()=>window.print();</script>
</body></html>`;

  const w = window.open('','_blank');
  if (!w) { showToast('Permite las ventanas emergentes para imprimir.','error'); return; }
  w.document.write(html);
  w.document.close();
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
// 1. Llena el Word con los datos del despacho
// 2. Envía el docx al microservicio en Render
// 3. Recibe el PDF y lo abre para imprimir
// ─────────────────────────────────────────

// URL del microservicio — actualiza esto cuando lo despliegues en Render
const CONVERTER_URL = 'https://innova-converter.onrender.com/convert';

async function generarYAbrirPDF(memo) {
  // ── Paso 1: generar el docx llenado en el navegador ──
  if (!window.JSZip) {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  const templateUrl = new URL('Entrega_de_materiales_OTC_2026.docx', window.location.href).href;
  const resp = await fetch(templateUrl);
  if (!resp.ok) throw new Error('No se encontró la plantilla Word en el repositorio.');
  const arrayBuffer = await resp.arrayBuffer();

  const zip    = await JSZip.loadAsync(arrayBuffer);
  const xmlStr = await zip.file('word/document.xml').async('string');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');
  const tables = xmlDoc.querySelectorAll('tbl');
  const ns     = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  function appendTextToCell(cell, value) {
    const paras = cell.querySelectorAll('p');
    const para  = paras[paras.length - 1] || paras[0];
    if (!para) return;
    const run = xmlDoc.createElementNS(ns, 'w:r');
    const t   = xmlDoc.createElementNS(ns, 'w:t');
    t.setAttribute('xml:space', 'preserve');
    t.textContent = ' ' + value;
    run.appendChild(t);
    para.appendChild(run);
  }

  function setCellText(cell, value) {
    cell.querySelectorAll('r').forEach(r => {
      r.querySelectorAll('t').forEach(t => { t.textContent = ''; });
    });
    const para = cell.querySelector('p');
    if (!para) return;
    const run = xmlDoc.createElementNS(ns, 'w:r');
    const t   = xmlDoc.createElementNS(ns, 'w:t');
    t.setAttribute('xml:space', 'preserve');
    t.textContent = String(value);
    run.appendChild(t);
    para.appendChild(run);
  }

  // Encabezado (Tabla 0)
  const HEADER_MAP_DOCX = {
    USUARIO_RESPONSABLE:    [3, 1],
    EMPRESA_CONTRATISTA:    [4, 0],
    INSTALADOR_RESPONSABLE: [4, 1],
    ENTREGADO_POR:          [5, 0],
    PLACA_VEHICULO:         [6, 1],
    FECHA_SOLICITUD:        [7, 0],
    FECHA_ENTREGA:          [7, 1],
  };
  const t0rows = tables[0].querySelectorAll('tr');
  for (const [campo, [fila, col]] of Object.entries(HEADER_MAP_DOCX)) {
    const valor = memo[campo];
    if (!valor || valor === '—') continue;
    const row   = t0rows[fila];
    if (!row) continue;
    const cells = row.querySelectorAll('tc');
    const cell  = cells[col];
    if (cell) appendTextToCell(cell, valor);
  }

  // Cantidades (Tabla 1)
  const SAP_ROW_MAP = {
    "221477":2,"213719":3,"328541":4,"352453":5,"352460":6,
    "352461":7,"352462":8,"353112":9,"354045":10,"354549":11,
    "200129":12,"355518":13,"338362":14,"219359":15,
    "328560":17,"243940":18,"337775":19,"337776":20,"337777":21,
    "210525":22,"212720":23,"214221":24,"301560":25,"245979":26,
    "355070":27,"244569":28,"353992":29,"211373":30,"211375":31,
    "353121":32,"355064":33,"338357":34,"353099":35,"353110":36,
    "353088":37,"200468":39,"200472":40,"200469":41,"200473":42,
    "213410":43,"214726":44,"214727":45,"352463":46,
    "219527":48,"221062":49,"200367":50,"282485":51,"350560":52,
    "212896":53,"350564":54,"221472":56,"200413":57,"211829":58,
    "213340":59,"222315":60,"353730":61,"338363":62,"338361":63,"338360":64,
  };
  const t1rows = tables[1].querySelectorAll('tr');
  for (const item of memo.MATERIALES) {
    const sap    = String(item.RESERVA).trim();
    const rowIdx = SAP_ROW_MAP[sap];
    if (rowIdx === undefined || !item.CANTIDAD) continue;
    const row   = t1rows[rowIdx];
    if (!row) continue;
    const cells = row.querySelectorAll('tc');
    if (cells[3]) setCellText(cells[3], String(item.CANTIDAD));
  }

  const serializer = new XMLSerializer();
  zip.file('word/document.xml', serializer.serializeToString(xmlDoc));
  const docxBlob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });

  // ── Paso 2: convertir a PDF via Render ──
  const docxBase64 = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(docxBlob);
  });

  const convResp = await fetch(CONVERTER_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ docx_base64: docxBase64 }),
  });

  if (!convResp.ok) {
    const err = await convResp.json().catch(() => ({}));
    throw new Error(err.error || `Error del servidor: ${convResp.status}`);
  }

  const { pdf_base64 } = await convResp.json();
  if (!pdf_base64) throw new Error('El servidor no devolvió el PDF.');

  // ── Paso 3: abrir PDF para imprimir ──
  const pdfBytes  = Uint8Array.from(atob(pdf_base64), c => c.charCodeAt(0));
  const pdfBlob   = new Blob([pdfBytes], { type: 'application/pdf' });
  const pdfUrl    = URL.createObjectURL(pdfBlob);
  const w         = window.open(pdfUrl, '_blank');
  if (!w) {
    // Fallback: descarga directa si el navegador bloquea el popup
    const a    = document.createElement('a');
    a.href     = pdfUrl;
    a.download = `despacho-${memo.FECHA_ENTREGA || 'sin-fecha'}.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
  setTimeout(() => URL.revokeObjectURL(pdfUrl), 10000);
}
