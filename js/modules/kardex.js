/**
 * kardex.js — Fase 1 (v3)
 * Control de materiales con códigos SAP y AX.
 *
 * Cambios:
 * - Modelo extendido: sapCode, axCode, name, unit, minStock, stock
 * - Formulario incluye campos SAP y AX
 * - Vista de inventario muestra nombre + códigos en texto secundario
 * - Búsqueda por nombre, SAP o AX
 * - Validación de duplicados por SAP o AX antes de crear
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
function safeStr(val, fb = '—') { return (val !== undefined && val !== null && String(val).trim() !== '') ? String(val).trim() : fb; }

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

// Filtra documentos sin datos reales (temporales)
function esValido(item) {
  const n = safeStr(item.name,'') || safeStr(item.nombre,'');
  const u = safeStr(item.unit,'') || safeStr(item.unidad,'');
  return n !== '' && u !== '';
}

function normalizeItem(raw) {
  return {
    ...raw,
    name:     safeStr(raw.name,'')    || safeStr(raw.nombre,''),
    unit:     safeStr(raw.unit,'')    || safeStr(raw.unidad,''),
    sapCode:  safeStr(raw.sapCode,''),
    axCode:   safeStr(raw.axCode,''),
    stock:    safeNum(raw.stock),
    minStock: safeNum(raw.minStock || raw.stockMinimo || 5),
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

    const qS      = query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'));
    const snapS   = await getDocs(qS);
    const ultimas = snapS.docs.slice(0, 5).map(d => ({ id: d.id, ...d.data() }));
    const alertas = items.filter(i => safeNum(i.stock) <= safeNum(i.minStock || 5));

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
                    <p class="text-sm font-medium text-gray-900 truncate">${safeStr(s.tecnicoNombre)}</p>
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
              const bs    = stock <= 0 ? 'background:#FEE2E2;color:#C62828' : 'background:#FEF3C7;color:#E65100';
              return `
              <div class="px-4 py-3 flex items-center justify-between gap-2">
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
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  const canEdit = ['admin','coordinadora'].includes(session.role);

  try {
    const snap  = await getDocs(collection(db, 'kardex/inventario/items'));
    const allItems = snap.docs.map(d => normalizeItem({ id: d.id, ...d.data() })).filter(esValido)
      .sort((a,b) => safeStr(a.name).localeCompare(safeStr(b.name)));

    content.innerHTML = `
      <div class="space-y-3">
        ${canEdit ? `
        <div class="flex gap-2 justify-end">
          ${session.role === 'admin' ? `
          <button id="btn-importar-excel"
            class="inline-flex items-center gap-2 text-sm font-medium px-3 py-2 rounded-lg border border-gray-300 text-gray-700 hover:bg-gray-50">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
              <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
            </svg>
            Importar Excel
          </button>` : ''}
          <button id="btn-nuevo-item"
            class="inline-flex items-center gap-2 text-white text-sm font-medium px-3 py-2 rounded-lg"
            style="background-color:#1B4F8A">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
            </svg>
            Agregar
          </button>
        </div>` : ''}

        <!-- Buscador -->
        <div class="relative">
          <svg class="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input id="inv-search" type="text" placeholder="Buscar por nombre, SAP o AX..."
            class="w-full border border-gray-300 rounded-xl pl-9 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>

        <div id="inv-tabla" class="bg-white rounded-xl border border-gray-200 overflow-hidden"></div>
      </div>`;

    function renderTabla(items) {
      const tabla = document.getElementById('inv-tabla');
      if (!tabla) return;
      if (items.length === 0) {
        tabla.innerHTML = `<div class="py-12 text-center text-sm text-gray-400">
          Sin resultados.${canEdit ? '<br>Agrega materiales con el botón de arriba.' : ''}
        </div>`;
        return;
      }
      tabla.innerHTML = `
        <div class="overflow-x-auto">
          <table class="w-full text-sm">
            <thead class="bg-gray-50 border-b border-gray-200">
              <tr>
                <th class="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Material</th>
                <th class="text-center px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide w-20">Stock</th>
                ${canEdit ? '<th class="px-4 py-3 w-20"></th>' : ''}
              </tr>
            </thead>
            <tbody class="divide-y divide-gray-100">
              ${items.map(item => itemRow(item, canEdit)).join('')}
            </tbody>
          </table>
        </div>`;

      if (canEdit) {
        tabla.querySelectorAll('[data-entrada]').forEach(b => {
          b.addEventListener('click', () => showFormEntrada(db, session, allItems.find(i => i.id === b.dataset.entrada)));
        });
        tabla.querySelectorAll('[data-edit]').forEach(b => {
          b.addEventListener('click', () => showFormItem(db, session, allItems.find(i => i.id === b.dataset.edit)));
        });
      }
    }

    renderTabla(allItems);

    // Búsqueda en tiempo real
    document.getElementById('inv-search')?.addEventListener('input', (e) => {
      const q = e.target.value.trim().toLowerCase();
      if (!q) { renderTabla(allItems); return; }
      const filtered = allItems.filter(i =>
        safeStr(i.name).toLowerCase().includes(q) ||
        safeStr(i.sapCode,'').toLowerCase().includes(q) ||
        safeStr(i.axCode,'').toLowerCase().includes(q)
      );
      renderTabla(filtered);
    });

    if (canEdit) {
      document.getElementById('btn-nuevo-item')?.addEventListener('click', () => showFormItem(db, session, null));
      document.getElementById('btn-importar-excel')?.addEventListener('click', () => showImportExcel(db, session, allItems, renderTabla));
    }

  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function itemRow(item, canEdit) {
  const stock  = safeNum(item.stock);
  const min    = safeNum(item.minStock || 5);
  const unit   = safeStr(item.unit, '');
  const name   = safeStr(item.name, '—');
  const sap    = safeStr(item.sapCode, '');
  const ax     = safeStr(item.axCode, '');

  const badgeStyle = stock <= 0
    ? 'background:#FEE2E2;color:#C62828'
    : stock <= min
    ? 'background:#FEF3C7;color:#E65100'
    : 'background:#DCFCE7;color:#166534';

  const codigos = [
    sap ? `<span class="font-mono">SAP ${sap}</span>` : '',
    ax  ? `<span class="font-mono">AX ${ax}</span>`   : '',
  ].filter(Boolean).join('<span class="mx-1 text-gray-300">·</span>');

  return `
    <tr class="hover:bg-gray-50">
      <td class="px-4 py-3">
        <p class="font-medium text-gray-900 leading-tight">${name}</p>
        ${codigos ? `<p class="text-xs text-gray-400 mt-0.5">${codigos}</p>` : ''}
      </td>
      <td class="px-4 py-3 text-center">
        <span class="text-sm font-bold px-2.5 py-1 rounded-full" style="${badgeStyle}">${stock}</span>
        <p class="text-xs text-gray-400 mt-0.5">${unit}</p>
      </td>
      ${canEdit ? `
      <td class="px-4 py-3">
        <div class="flex items-center justify-end gap-1">
          <button data-entrada="${item.id}" title="Registrar entrada"
            class="p-1.5 rounded text-gray-400 hover:text-green-600 transition-colors">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M12 5v14M5 12l7 7 7-7"/>
            </svg>
          </button>
          <button data-edit="${item.id}" title="Editar"
            class="p-1.5 rounded text-gray-400 hover:text-blue-500 transition-colors">
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
          <input id="fi-sap" type="text" value="${safeStr(item?.sapCode,'')}"
            placeholder="Ej. 221477"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1">Código AX</label>
          <input id="fi-ax" type="text" value="${safeStr(item?.axCode,'')}"
            placeholder="Ej. 50203"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2">
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1">Unidad</label>
          <input id="fi-unit" type="text" value="${safeStr(item?.unit||item?.unidad,'')}"
            placeholder="unidad, m, kg"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1">Stock mínimo</label>
          <input id="fi-min" type="number" min="0" value="${safeNum(item?.minStock||item?.stockMinimo) || 5}"
            class="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
      </div>
      ${esNuevo ? `
      <div class="bg-blue-50 border border-blue-100 rounded-lg p-2.5">
        <p class="text-xs font-medium text-blue-800 mb-1">Stock inicial <span class="font-normal text-blue-500">(se registra como entrada)</span></p>
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
    const name    = ov.querySelector('#fi-name').value.trim();
    const sapCode = ov.querySelector('#fi-sap').value.trim();
    const axCode  = ov.querySelector('#fi-ax').value.trim();
    const unit    = ov.querySelector('#fi-unit').value.trim();
    const minStock = safeNum(ov.querySelector('#fi-min').value);
    const errEl   = ov.querySelector('#fi-err');
    const btn     = ov.querySelector('#fi-submit');

    errEl.classList.add('hidden');
    if (!name) { errEl.textContent = 'El nombre es obligatorio.'; errEl.classList.remove('hidden'); return; }
    if (!unit) { errEl.textContent = 'La unidad de medida es obligatoria.'; errEl.classList.remove('hidden'); return; }

    btn.disabled = true; btn.textContent = 'Verificando...';

    try {
      // Verificar duplicados por SAP o AX (solo al crear, o si el código cambió)
      if (sapCode || axCode) {
        const snapAll = await getDocs(collection(db, 'kardex/inventario/items'));
        const existing = snapAll.docs.map(d => ({ id: d.id, ...d.data() })).filter(esValido);

        for (const ex of existing) {
          if (item && ex.id === item.id) continue; // excluir el mismo item al editar
          if (sapCode && safeStr(ex.sapCode,'') === sapCode) {
            errEl.textContent = `El código SAP ${sapCode} ya existe en el material: ${safeStr(ex.name)}`;
            errEl.classList.remove('hidden');
            btn.disabled = false; btn.textContent = esNuevo ? 'Agregar' : 'Guardar cambios';
            return;
          }
          if (axCode && safeStr(ex.axCode,'') === axCode) {
            errEl.textContent = `El código AX ${axCode} ya existe en el material: ${safeStr(ex.name)}`;
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
          name, sapCode, axCode, unit,
          stock: stockInicial,
          minStock,
          creadoEn:   serverTimestamp(),
          creadoPor:  session.uid,
        });
        if (stockInicial > 0) {
          await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
            itemId: ref.id, itemNombre: name,
            tipo: 'entrada_inicial', cantidad: stockInicial,
            motivo: 'Stock inicial al crear el material',
            stockAntes: 0, stockDespues: stockInicial,
            fecha: serverTimestamp(),
            registradoPor: session.uid,
            registradoPorNombre: session.displayName,
          });
        }
      } else {
        await updateDoc(doc(db, 'kardex/inventario/items', item.id), {
          name, sapCode, axCode, unit, minStock,
        });
      }

      ov.remove();
      showToast(`Material ${esNuevo ? 'agregado' : 'actualizado'} correctamente.`, 'success');
      await showInventario(db, session);
    } catch(e) {
      errEl.textContent = 'Error al guardar. Intenta de nuevo.';
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
          ${item?.sapCode ? `SAP ${item.sapCode}` : ''}
          ${item?.sapCode && item?.axCode ? ' · ' : ''}
          ${item?.axCode  ? `AX ${item.axCode}`  : ''}
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
        <input id="fe-motivo" type="text" placeholder="Ej. Compra de materiales, Reposición"
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
    if (cant <= 0) { errEl.textContent = 'Ingresa una cantidad válida mayor a 0.'; errEl.classList.remove('hidden'); return; }
    btn.disabled = true; btn.textContent = 'Registrando...';
    try {
      const stockAntes   = safeNum(item?.stock);
      const stockDespues = stockAntes + cant;
      await updateDoc(doc(db, 'kardex/inventario/items', item.id), { stock: increment(cant) });
      await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
        itemId: item.id, itemNombre: safeStr(item?.name),
        tipo: 'entrada', cantidad: cant, motivo,
        stockAntes, stockDespues,
        fecha: serverTimestamp(),
        registradoPor: session.uid, registradoPorNombre: session.displayName,
      });
      ov.remove();
      showToast(`Entrada de ${cant} ${safeStr(item?.unit,'')} registrada.`, 'success');
      await showInventario(db, session);
    } catch(e) {
      errEl.textContent = 'Error al registrar.';
      errEl.classList.remove('hidden');
      btn.disabled = false; btn.textContent = 'Registrar entrada';
      console.error(e);
    }
  };
}

// ─────────────────────────────────────────
// FORM — NUEVA SALIDA
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

  const tecnicos = snapU.docs
    .map(d => d.data())
    .filter(u => u.role === 'campo')
    .sort((a,b) => safeStr(a.displayName).localeCompare(safeStr(b.displayName)));

  let sel = [];

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Nueva salida de materiales</h2>
      <button id="fs-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div class="px-5 py-4 space-y-4">
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Técnico que recibe</label>
        <select id="fs-tec" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400">
          <option value="">Seleccionar técnico...</option>
          ${tecnicos.map(t => `<option value="${t.uid}" data-nombre="${safeStr(t.displayName)}">${safeStr(t.displayName)}</option>`).join('')}
        </select>
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Motivo / Trabajo</label>
        <input id="fs-motivo" type="text" placeholder="Ej. Cambio de medidor, Servicio nuevo"
          class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Agregar materiales</label>
        <div class="flex gap-2">
          <select id="fs-item" class="flex-1 border border-gray-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 min-w-0">
            <option value="">Seleccionar material...</option>
            ${items.map(i => `
              <option value="${i.id}"
                data-nombre="${safeStr(i.name)}"
                data-unit="${safeStr(i.unit,'')}"
                data-stock="${safeNum(i.stock)}"
                data-sap="${safeStr(i.sapCode,'')}"
              >${safeStr(i.name)} (${safeNum(i.stock)} ${safeStr(i.unit,'')})</option>`).join('')}
          </select>
          <input id="fs-cant" type="number" min="1" placeholder="Cant."
            class="w-20 border border-gray-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 text-center shrink-0"/>
          <button id="fs-add"
            class="px-3 py-2.5 text-white rounded-lg text-sm font-bold shrink-0" style="background-color:#2196F3">+</button>
        </div>
      </div>
      <div id="fs-lista" class="space-y-2"></div>
      <div id="fs-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fs-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cancelar</button>
      <button id="fs-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">Registrar salida</button>
    </div>`);

  function renderLista() {
    const lista = ov.querySelector('#fs-lista');
    lista.innerHTML = sel.length === 0
      ? '<p class="text-xs text-gray-400 text-center py-2">Sin materiales agregados</p>'
      : sel.map((s,i) => `
        <div class="flex items-center justify-between bg-gray-50 rounded-lg px-3 py-2 gap-2">
          <div class="min-w-0">
            <p class="text-sm font-medium text-gray-900 truncate">${s.nombre}</p>
            <p class="text-xs text-gray-500">${s.cantidad} ${s.unit}</p>
          </div>
          <button data-ri="${i}" class="text-gray-400 hover:text-red-500 shrink-0 text-lg leading-none">✕</button>
        </div>`).join('');
    lista.querySelectorAll('[data-ri]').forEach(b => {
      b.onclick = () => { sel.splice(parseInt(b.dataset.ri),1); renderLista(); };
    });
  }
  renderLista();

  ov.querySelector('#fs-close').onclick = ov.querySelector('#fs-cancel').onclick = () => ov.remove();

  ov.querySelector('#fs-add').onclick = () => {
    const selEl = ov.querySelector('#fs-item');
    const opt   = selEl.options[selEl.selectedIndex];
    const cant  = safeNum(ov.querySelector('#fs-cant').value);
    const errEl = ov.querySelector('#fs-err');
    errEl.classList.add('hidden');
    if (!selEl.value) { errEl.textContent = 'Selecciona un material.'; errEl.classList.remove('hidden'); return; }
    if (cant <= 0)    { errEl.textContent = 'Cantidad inválida.';       errEl.classList.remove('hidden'); return; }
    const disponible = safeNum(opt.dataset.stock);
    if (cant > disponible) {
      errEl.textContent = `Stock insuficiente. Disponible: ${disponible} ${opt.dataset.unit}`;
      errEl.classList.remove('hidden'); return;
    }
    const ex = sel.findIndex(s => s.itemId === selEl.value);
    if (ex >= 0) sel[ex].cantidad += cant;
    else sel.push({
      itemId:  selEl.value,
      nombre:  safeStr(opt.dataset.nombre),
      unit:    safeStr(opt.dataset.unit,''),
      cantidad: cant,
    });
    selEl.value = ''; ov.querySelector('#fs-cant').value = '';
    renderLista();
  };

  ov.querySelector('#fs-submit').onclick = async () => {
    const tecSel = ov.querySelector('#fs-tec');
    const tecUid = tecSel.value;
    const tecNom = safeStr(tecSel.options[tecSel.selectedIndex]?.dataset.nombre);
    const motivo = ov.querySelector('#fs-motivo').value.trim();
    const errEl  = ov.querySelector('#fs-err');
    const btn    = ov.querySelector('#fs-submit');
    errEl.classList.add('hidden');
    if (!tecUid)       { errEl.textContent = 'Selecciona el técnico.';           errEl.classList.remove('hidden'); return; }
    if (sel.length===0){ errEl.textContent = 'Agrega al menos un material.';     errEl.classList.remove('hidden'); return; }
    btn.disabled = true; btn.textContent = 'Registrando...';
    try {
      const ref = await addDoc(collection(db, 'kardex/movimientos/salidas'), {
        tecnicoUid: tecUid, tecnicoNombre: tecNom, motivo,
        items: sel.map(s => ({ itemId:s.itemId, nombre:s.nombre, unit:s.unit, cantidad:s.cantidad })),
        fecha: serverTimestamp(),
        registradoPor: session.uid, registradoPorNombre: session.displayName,
      });
      for (const s of sel) {
        await updateDoc(doc(db,'kardex/inventario/items',s.itemId), { stock: increment(-s.cantidad) });
      }
      ov.remove();
      showToast('Salida registrada correctamente.','success');
      showMemo({ id:ref.id, tecnicoNombre:tecNom, motivo, items:sel, registradoPorNombre:session.displayName, fecha:new Date() });
      await showDashboard(db, session);
    } catch(e) {
      errEl.textContent = 'Error al registrar.'; errEl.classList.remove('hidden');
      btn.disabled = false; btn.textContent = 'Registrar salida'; console.error(e);
    }
  };
}

// ─────────────────────────────────────────
// HISTORIAL
// ─────────────────────────────────────────
async function showHistorial(db, session) {
  setTab('historial');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  try {
    let q;
    if (session.role === 'campo') {
      q = query(collection(db,'kardex/movimientos/salidas'), where('tecnicoUid','==',session.uid), orderBy('fecha','desc'));
    } else {
      q = query(collection(db,'kardex/movimientos/salidas'), orderBy('fecha','desc'));
    }
    const snap    = await getDocs(q);
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
                  <p class="text-sm font-semibold text-gray-900">${safeStr(s.tecnicoNombre)}</p>
                  <p class="text-xs text-gray-400">${fmtDate(s.fecha)} · ${safeStr(s.registradoPorNombre)}</p>
                </div>
                <button data-smemo="${s.id}"
                  class="text-xs font-medium px-2 py-1 rounded-lg shrink-0" style="color:#1B4F8A;background:#EFF6FF">Memo</button>
              </div>
              ${s.motivo ? `<p class="text-xs text-gray-500 mt-1">${s.motivo}</p>` : ''}
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
        if (s) showMemo({
          id: s.id,
          tecnicoNombre: safeStr(s.tecnicoNombre),
          motivo: s.motivo,
          items: s.items || [],
          registradoPorNombre: safeStr(s.registradoPorNombre),
          fecha: s.fecha?.toDate?.() || new Date(),
        });
      };
    });
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ─────────────────────────────────────────
// MEMO
// ─────────────────────────────────────────
function showMemo(s) {
  const fecha    = s.fecha instanceof Date ? s.fecha : new Date();
  const fechaStr = fecha.toLocaleDateString('es-SV', { year:'numeric', month:'long', day:'numeric' });

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Memo de salida</h2>
      <button id="fm-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div id="memo-body" class="px-6 py-5 space-y-4">
      <div class="text-center border-b border-gray-200 pb-4">
        <p class="font-bold text-lg" style="color:#1B4F8A">INNOVA STC</p>
        <p class="text-sm text-gray-500">Servicios Técnicos y Comerciales</p>
        <p class="text-xs text-gray-400 mt-1 uppercase tracking-wide">Memo de Salida de Materiales</p>
      </div>
      <div class="grid grid-cols-2 gap-3 text-sm">
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">Fecha</p><p class="font-medium text-gray-900 mt-0.5">${fechaStr}</p></div>
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">N° Ref.</p><p class="font-medium font-mono text-gray-900 mt-0.5">${safeStr(s.id).slice(-6).toUpperCase()}</p></div>
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">Técnico</p><p class="font-medium text-gray-900 mt-0.5">${safeStr(s.tecnicoNombre)}</p></div>
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">Entregado por</p><p class="font-medium text-gray-900 mt-0.5">${safeStr(s.registradoPorNombre)}</p></div>
      </div>
      ${s.motivo ? `<div><p class="text-xs text-gray-400 uppercase tracking-wide">Motivo / Trabajo</p><p class="text-sm font-medium text-gray-900 mt-0.5">${s.motivo}</p></div>` : ''}
      <div>
        <p class="text-xs text-gray-400 uppercase tracking-wide mb-2">Materiales entregados</p>
        <table class="w-full text-sm border border-gray-200 rounded-lg overflow-hidden">
          <thead class="bg-gray-50">
            <tr>
              <th class="text-left px-3 py-2 text-xs font-semibold text-gray-600">Material</th>
              <th class="text-center px-3 py-2 text-xs font-semibold text-gray-600">Cant.</th>
              <th class="text-left px-3 py-2 text-xs font-semibold text-gray-600">Unidad</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-gray-100">
            ${(s.items||[]).map(i => `
              <tr>
                <td class="px-3 py-2 text-gray-900">${safeStr(i.nombre||i.name)}</td>
                <td class="px-3 py-2 text-center font-semibold text-gray-900">${safeNum(i.cantidad)}</td>
                <td class="px-3 py-2 text-gray-500">${safeStr(i.unit||i.unidad,'')}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
      <div class="grid grid-cols-2 gap-6 pt-4">
        <div class="text-center">
          <div class="border-b border-gray-400 mb-1 h-10"></div>
          <p class="text-xs text-gray-500">Entregado por</p>
          <p class="text-xs font-medium text-gray-700">${safeStr(s.registradoPorNombre)}</p>
        </div>
        <div class="text-center">
          <div class="border-b border-gray-400 mb-1 h-10"></div>
          <p class="text-xs text-gray-500">Recibido por</p>
          <p class="text-xs font-medium text-gray-700">${safeStr(s.tecnicoNombre)}</p>
        </div>
      </div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fm-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cerrar</button>
      <button id="fm-print"  class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">🖨️ Imprimir</button>
    </div>`);

  ov.querySelector('#fm-close').onclick = ov.querySelector('#fm-cancel').onclick = () => ov.remove();
  ov.querySelector('#fm-print').onclick = () => {
    const body = document.getElementById('memo-body').innerHTML;
    const w = window.open('','_blank');
    w.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Memo INNOVA STC</title>
      <style>body{font-family:Arial,sans-serif;font-size:13px;color:#111;margin:24px}
      table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:6px 10px;text-align:left}
      th{background:#f5f5f5;font-size:11px}@media print{body{margin:10mm}}</style>
      </head><body>${body}</body></html>`);
    w.document.close(); w.print();
  };
}

// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────
function mkOverlay(inner) {
  const ov = document.createElement('div');
  ov.className = 'fixed inset-0 bg-black/50 flex items-start justify-center z-50 p-4 overflow-y-auto';
  ov.innerHTML = `<div class="bg-white rounded-2xl shadow-xl w-full max-w-md my-4" style="max-height:90vh;overflow-y:auto;">${inner}</div>`;
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
    return d.toLocaleDateString('es-SV',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
  } catch { return '—'; }
}

// ─────────────────────────────────────────
// IMPORTAR DESDE EXCEL
// Solo admin. Modal integrado en el Kardex.
// ─────────────────────────────────────────
async function showImportExcel(db, session, existingItems, refreshFn) {
  // Cargar SheetJS dinámicamente si no está disponible
  if (!window.XLSX) {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js';
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  let toImport = [];
  let toSkip   = [];

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Importar desde Excel</h2>
      <button id="ix-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>

    <!-- Paso A: seleccionar archivo -->
    <div id="ix-step-file" class="px-5 py-4 space-y-3">
      <p class="text-sm text-gray-500">Selecciona el archivo Excel con la lista de materiales.</p>
      <p class="text-xs text-gray-400">Columnas requeridas: <strong>SAP</strong>, <strong>AX</strong>, <strong>Nombre material</strong></p>
      <label class="flex flex-col items-center justify-center border-2 border-dashed border-gray-300 rounded-xl p-5 cursor-pointer hover:border-blue-400 transition-colors">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#9CA3AF" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" class="mb-2">
          <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
          <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
        <p class="text-sm text-gray-600 font-medium">Toca para seleccionar el Excel</p>
        <p class="text-xs text-gray-400 mt-0.5">.xlsx</p>
        <input id="ix-file" type="file" accept=".xlsx,.xls" class="hidden"/>
      </label>
      <div id="ix-file-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2"></div>
    </div>

    <!-- Paso B: previsualización -->
    <div id="ix-step-preview" class="hidden px-5 py-4 space-y-3">
      <div id="ix-summary" class="grid grid-cols-3 gap-2"></div>
      <div id="ix-tabla" class="rounded-lg border border-gray-200 overflow-hidden max-h-52 overflow-y-auto text-xs"></div>
      <div id="ix-skipped" class="hidden"></div>
      <div id="ix-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2"></div>
      <div class="flex gap-3 pt-1">
        <button id="ix-back" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">← Volver</button>
        <button id="ix-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background:#1B4F8A">Importar</button>
      </div>
    </div>

    <!-- Paso C: resultado -->
    <div id="ix-step-result" class="hidden px-5 py-6 text-center space-y-3">
      <div class="w-14 h-14 rounded-full flex items-center justify-center mx-auto" style="background:#DCFCE7">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#166534" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="20 6 9 17 4 12"/>
        </svg>
      </div>
      <p class="font-semibold text-gray-900">Importación completada</p>
      <div id="ix-result-txt" class="space-y-1"></div>
      <button id="ix-done" class="text-white font-medium rounded-lg px-6 py-2.5 text-sm mt-2" style="background:#1B4F8A">Cerrar</button>
    </div>`);

  ov.querySelector('#ix-close').onclick = () => ov.remove();

  // ── Leer archivo ──
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

      // Mapear columnas flexible
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

      // Verificar duplicados contra items existentes
      const existingSap = new Set(existingItems.map(i => safeStr(i.sapCode,'')).filter(Boolean));
      const existingAx  = new Set(existingItems.map(i => safeStr(i.axCode,'') ).filter(Boolean));

      toImport = []; toSkip = [];
      for (const m of materiales) {
        const dupSap = m.sapCode && existingSap.has(m.sapCode);
        const dupAx  = m.axCode  && existingAx.has(m.axCode);
        if (dupSap || dupAx) toSkip.push({ ...m, reason: dupSap ? `SAP ${m.sapCode} ya existe` : `AX ${m.axCode} ya existe` });
        else toImport.push(m);
      }

      // Mostrar previsualización
      ov.querySelector('#ix-step-file').classList.add('hidden');
      ov.querySelector('#ix-step-preview').classList.remove('hidden');

      // Stats
      ov.querySelector('#ix-summary').innerHTML = [
        { v: toImport.length, l: 'A importar', c: '#2196F3' },
        { v: toSkip.length,   l: 'Saltados',   c: '#E65100' },
        { v: materiales.length, l: 'Total',    c: '#6B7280' },
      ].map(s => `<div class="bg-gray-50 rounded-xl p-2 text-center border border-gray-200">
        <p class="text-xl font-bold" style="color:${s.c}">${s.v}</p>
        <p class="text-xs text-gray-500">${s.l}</p>
      </div>`).join('');

      // Tabla
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

      // Saltados
      if (toSkip.length > 0) {
        const sk = ov.querySelector('#ix-skipped');
        sk.classList.remove('hidden');
        sk.innerHTML = `<div class="bg-orange-50 border border-orange-200 rounded-xl p-3">
          <p class="text-xs font-semibold text-orange-800 mb-1">⚠️ ${toSkip.length} serán saltados (ya existen)</p>
          ${toSkip.slice(0,5).map(m => `<p class="text-xs text-orange-600 truncate">${m.name} — ${m.reason}</p>`).join('')}
          ${toSkip.length > 5 ? `<p class="text-xs text-orange-400">...y ${toSkip.length-5} más</p>` : ''}
        </div>`;
      }

    } catch(e) {
      errEl.textContent = e.message || 'Error al leer el archivo.';
      errEl.classList.remove('hidden');
    }
  });

  // ── Volver ──
  ov.querySelector('#ix-back').onclick = () => {
    ov.querySelector('#ix-step-preview').classList.add('hidden');
    ov.querySelector('#ix-step-file').classList.remove('hidden');
    ov.querySelector('#ix-file').value = '';
  };

  // ── Importar ──
  ov.querySelector('#ix-submit').onclick = async () => {
    if (toImport.length === 0) {
      ov.querySelector('#ix-err').textContent = 'No hay materiales nuevos para importar.';
      ov.querySelector('#ix-err').classList.remove('hidden');
      return;
    }

    const btn = ov.querySelector('#ix-submit');
    btn.disabled = true; btn.textContent = 'Importando...';

    let ok = 0, errCount = 0;
    for (const m of toImport) {
      try {
        await addDoc(collection(db, 'kardex/inventario/items'), {
          name:      m.name,
          sapCode:   m.sapCode,
          axCode:    m.axCode,
          unit:      'unidad',
          stock:     isNaN(m.stock) ? 0 : m.stock,
          minStock:  0,
          creadoEn:  serverTimestamp(),
          creadoPor: session.uid,
        });
        ok++;
      } catch(e) { errCount++; console.error(e); }
      btn.textContent = `Importando ${ok+errCount}/${toImport.length}...`;
    }

    ov.querySelector('#ix-step-preview').classList.add('hidden');
    ov.querySelector('#ix-step-result').classList.remove('hidden');
    ov.querySelector('#ix-result-txt').innerHTML = `
      <p class="text-sm text-gray-700"><strong class="text-green-700">${ok}</strong> material(es) importados.</p>
      ${toSkip.length > 0 ? `<p class="text-sm text-gray-500">${toSkip.length} saltados por duplicado.</p>` : ''}
      ${errCount > 0 ? `<p class="text-sm text-red-600">${errCount} errores.</p>` : ''}
      <p class="text-xs text-gray-400 mt-1">Revisa la unidad de cada material importado.</p>`;
  };

  // ── Cerrar tras resultado ──
  ov.querySelector('#ix-done').onclick = async () => {
    ov.remove();
    await showInventario(db, session);
  };
}
