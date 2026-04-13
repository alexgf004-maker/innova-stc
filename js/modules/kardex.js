/**
 * kardex.js — Fase 1 (revisado)
 * Control de materiales mediante movimientos.
 *
 * Cambios respecto a versión anterior:
 * - Badge de stock muestra solo el número, unidad aparte
 * - Conversión segura a número en todos los campos numéricos
 * - Filtrado de documentos sin nombre/unidad (documentos temporales)
 * - Formulario de material solo crea en catálogo, stock inicial como entrada controlada
 * - Botón de ajuste renombrado a "Entrada" para reflejar el flujo real
 * - Dashboard prioriza acciones sobre inventario
 * - Nunca aparece "undefined" en interfaz
 */

import {
  collection, doc, getDocs, addDoc, updateDoc,
  query, orderBy, where, serverTimestamp, increment
} from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';

import { showToast } from '../ui.js';

// ─────────────────────────────────────────
// HELPERS SEGUROS
// ─────────────────────────────────────────
function safeNum(val) {
  const n = Number(val);
  return isNaN(n) ? 0 : n;
}

function safeStr(val, fallback = '—') {
  return (val !== undefined && val !== null && val !== '') ? String(val) : fallback;
}

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
    b.style.color = a ? '#1B4F8A' : '#6B7280';
    b.style.boxShadow = a ? '0 1px 3px rgba(0,0,0,0.1)' : 'none';
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

// ─────────────────────────────────────────
// FILTRO: excluye documentos sin datos reales
// ─────────────────────────────────────────
function esItemValido(item) {
  return item.nombre && item.nombre !== '' && item.unidad && item.unidad !== '';
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
    const items = snap.docs.map(d => ({ id: d.id, ...d.data() })).filter(esItemValido);

    const total     = items.length;
    const sinStock  = items.filter(i => safeNum(i.stock) <= 0).length;
    const stockBajo = items.filter(i => safeNum(i.stock) > 0 && safeNum(i.stock) <= safeNum(i.stockMinimo || 5)).length;

    const qS     = query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'));
    const snapS  = await getDocs(qS);
    const ultimas = snapS.docs.slice(0, 5).map(d => ({ id: d.id, ...d.data() }));

    const alertas = items.filter(i => safeNum(i.stock) <= safeNum(i.stockMinimo || 5));

    content.innerHTML = `
      <div class="space-y-4">

        <!-- Acciones rápidas -->
        <div class="grid grid-cols-2 gap-3">
          ${statCard(total, 'Materiales', '#2196F3')}
          ${statCard(sinStock, 'Sin stock', '#C62828')}
        </div>

        <!-- Últimas salidas -->
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

        <!-- Alertas de stock -->
        ${alertas.length > 0 ? `
        <div class="bg-white rounded-xl border border-orange-200 overflow-hidden">
          <div class="px-4 py-3 border-b border-orange-100">
            <p class="font-semibold text-sm text-orange-800">⚠️ Atención al inventario</p>
          </div>
          <div class="divide-y divide-gray-100">
            ${alertas.map(i => {
              const stock = safeNum(i.stock);
              const badge = stock <= 0
                ? 'background:#FEE2E2;color:#C62828'
                : 'background:#FEF3C7;color:#E65100';
              return `
              <div class="px-4 py-3 flex items-center justify-between">
                <div>
                  <p class="text-sm text-gray-900">${safeStr(i.nombre)}</p>
                  <p class="text-xs text-gray-400">${safeStr(i.unidad)}</p>
                </div>
                <span class="text-xs font-semibold px-2 py-1 rounded-full" style="${badge}">
                  ${stock <= 0 ? 'Sin stock' : stock}
                </span>
              </div>`;
            }).join('')}
          </div>
        </div>` : ''}

      </div>`;

    // Rebind tabs dentro del dashboard
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
  return `
    <div class="bg-white rounded-xl border border-gray-200 p-4 text-center">
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

  try {
    const snap  = await getDocs(collection(db, 'kardex/inventario/items'));
    const items = snap.docs
      .map(d => ({ id: d.id, ...d.data() }))
      .filter(esItemValido)
      .sort((a, b) => safeStr(a.nombre).localeCompare(safeStr(b.nombre)));

    content.innerHTML = `
      <div class="space-y-3">
        ${canEdit ? `
        <div class="flex justify-end">
          <button id="btn-nuevo-item"
            class="inline-flex items-center gap-2 text-white text-sm font-medium px-3 py-2 rounded-lg"
            style="background-color:#1B4F8A">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
            </svg>
            Agregar material
          </button>
        </div>` : ''}

        <div class="bg-white rounded-xl border border-gray-200 overflow-hidden">
          ${items.length === 0
            ? `<div class="py-12 text-center text-sm text-gray-400">
                No hay materiales en el catálogo.
                ${canEdit ? '<br>Agrega el primero con el botón de arriba.' : ''}
               </div>`
            : `<div class="overflow-x-auto">
                <table class="w-full text-sm">
                  <thead class="bg-gray-50 border-b border-gray-200">
                    <tr>
                      <th class="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Material</th>
                      <th class="text-center px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Stock</th>
                      ${canEdit ? '<th class="px-4 py-3 w-24"></th>' : ''}
                    </tr>
                  </thead>
                  <tbody class="divide-y divide-gray-100">
                    ${items.map(item => itemRow(item, canEdit)).join('')}
                  </tbody>
                </table>
               </div>`
          }
        </div>
      </div>`;

    if (canEdit) {
      document.getElementById('btn-nuevo-item')?.addEventListener('click', () => showFormItem(db, session, null));
      content.querySelectorAll('[data-entrada]').forEach(b => {
        b.addEventListener('click', () => showFormEntrada(db, session, items.find(i => i.id === b.dataset.entrada)));
      });
      content.querySelectorAll('[data-edit]').forEach(b => {
        b.addEventListener('click', () => showFormItem(db, session, items.find(i => i.id === b.dataset.edit)));
      });
    }
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function itemRow(item, canEdit) {
  const stock = safeNum(item.stock);
  const min   = safeNum(item.stockMinimo || 5);
  const unidad = safeStr(item.unidad, '');
  const nombre = safeStr(item.nombre, '—');

  // Color del badge: solo el número, sin unidad
  const badgeStyle = stock <= 0
    ? 'background:#FEE2E2;color:#C62828'
    : stock <= min
    ? 'background:#FEF3C7;color:#E65100'
    : 'background:#DCFCE7;color:#166534';

  return `
    <tr class="hover:bg-gray-50">
      <td class="px-4 py-3">
        <p class="font-medium text-gray-900">${nombre}</p>
        <p class="text-xs text-gray-400">${unidad}</p>
      </td>
      <td class="px-4 py-3 text-center">
        <span class="text-sm font-bold px-2.5 py-1 rounded-full" style="${badgeStyle}">${stock}</span>
        <p class="text-xs text-gray-400 mt-0.5">${unidad}</p>
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
          <button data-edit="${item.id}" title="Editar datos del material"
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
// FORM — AGREGAR / EDITAR MATERIAL (catálogo)
// Solo modifica nombre, unidad y stock mínimo.
// El stock inicial se registra como entrada controlada.
// ─────────────────────────────────────────
function showFormItem(db, session, item) {
  const esNuevo = !item;
  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">${esNuevo ? 'Nuevo material' : 'Editar material'}</h2>
      <button id="fi-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div class="px-5 py-4 space-y-4">
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Nombre del material</label>
        <input id="fi-nombre" type="text" value="${safeStr(item?.nombre,'')}"
          placeholder="Ej. Medidor monofásico 120V"
          class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
      </div>
      <div class="grid grid-cols-2 gap-3">
        <div>
          <label class="block text-sm font-medium text-gray-700 mb-1.5">Unidad de medida</label>
          <input id="fi-unidad" type="text" value="${safeStr(item?.unidad,'')}"
            placeholder="unidad, m, kg, rollo"
            class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
        <div>
          <label class="block text-sm font-medium text-gray-700 mb-1.5">Alerta stock mínimo</label>
          <input id="fi-min" type="number" min="0" value="${safeNum(item?.stockMinimo) || 5}"
            class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
        </div>
      </div>
      ${esNuevo ? `
      <div class="bg-blue-50 border border-blue-100 rounded-xl p-3">
        <p class="text-xs font-medium text-blue-800 mb-1">Stock inicial</p>
        <p class="text-xs text-blue-600 mb-2">Si ya tienes unidades en bodega, indica la cantidad. Esto se registrará como una entrada inicial.</p>
        <input id="fi-stock" type="number" min="0" value="0" placeholder="0"
          class="w-full border border-blue-200 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 bg-white"/>
      </div>` : `
      <div class="bg-gray-50 border border-gray-200 rounded-xl p-3">
        <p class="text-xs text-gray-500">Para modificar el stock usa el botón de <strong>entrada</strong> en la lista de inventario.</p>
      </div>`}
      <div id="fi-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fi-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cancelar</button>
      <button id="fi-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">
        ${esNuevo ? 'Agregar' : 'Guardar cambios'}
      </button>
    </div>`);

  ov.querySelector('#fi-close').onclick = ov.querySelector('#fi-cancel').onclick = () => ov.remove();
  ov.querySelector('#fi-submit').onclick = async () => {
    const nombre = ov.querySelector('#fi-nombre').value.trim();
    const unidad = ov.querySelector('#fi-unidad').value.trim();
    const minimo = safeNum(ov.querySelector('#fi-min').value);
    const errEl  = ov.querySelector('#fi-err');
    const btn    = ov.querySelector('#fi-submit');

    errEl.classList.add('hidden');
    if (!nombre) { errEl.textContent = 'El nombre es obligatorio.'; errEl.classList.remove('hidden'); return; }
    if (!unidad) { errEl.textContent = 'La unidad de medida es obligatoria.'; errEl.classList.remove('hidden'); return; }

    btn.disabled = true; btn.textContent = 'Guardando...';
    try {
      if (esNuevo) {
        const stockInicial = safeNum(ov.querySelector('#fi-stock').value);
        const ref = await addDoc(collection(db, 'kardex/inventario/items'), {
          nombre, unidad,
          stock: stockInicial,
          stockMinimo: minimo,
          creadoEn: serverTimestamp(),
          creadoPor: session.uid,
        });
        // Registrar entrada inicial si hay stock
        if (stockInicial > 0) {
          await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
            itemId: ref.id, itemNombre: nombre,
            tipo: 'entrada_inicial',
            cantidad: stockInicial,
            motivo: 'Stock inicial al crear el material',
            stockAntes: 0, stockDespues: stockInicial,
            fecha: serverTimestamp(),
            registradoPor: session.uid,
            registradoPorNombre: session.displayName,
          });
        }
      } else {
        // Solo actualiza datos del catálogo, no el stock
        await updateDoc(doc(db, 'kardex/inventario/items', item.id), {
          nombre, unidad, stockMinimo: minimo,
        });
      }
      ov.remove();
      showToast(`Material ${esNuevo ? 'agregado' : 'actualizado'} correctamente.`, 'success');
      await showInventario(db, session);
    } catch(e) {
      errEl.textContent = 'Error al guardar. Intenta de nuevo.';
      errEl.classList.remove('hidden');
      btn.disabled = false;
      btn.textContent = esNuevo ? 'Agregar' : 'Guardar cambios';
      console.error(e);
    }
  };
}

// ─────────────────────────────────────────
// FORM — ENTRADA DE MATERIAL
// Aumenta el stock mediante un movimiento registrado.
// ─────────────────────────────────────────
function showFormEntrada(db, session, item) {
  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Registrar entrada</h2>
      <button id="fe-close" class="text-gray-400 hover:text-gray-700">✕</button>
    </div>
    <div class="px-5 py-4 space-y-4">
      <div class="bg-gray-50 rounded-xl p-3">
        <p class="text-sm font-medium text-gray-900">${safeStr(item?.nombre)}</p>
        <p class="text-xs text-gray-500 mt-0.5">Stock actual: <strong>${safeNum(item?.stock)} ${safeStr(item?.unidad,'')}</strong></p>
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Cantidad que ingresa</label>
        <input id="fe-cant" type="number" min="1" placeholder="0"
          class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 text-center text-lg font-semibold"/>
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-700 mb-1.5">Motivo de entrada</label>
        <input id="fe-motivo" type="text" placeholder="Ej. Compra de materiales, Reposición"
          class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/>
      </div>
      <div id="fe-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fe-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cancelar</button>
      <button id="fe-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#2E7D32">
        Registrar entrada
      </button>
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

      await updateDoc(doc(db, 'kardex/inventario/items', item.id), {
        stock: increment(cant),
      });

      await addDoc(collection(db, 'kardex/movimientos/ajustes'), {
        itemId: item.id, itemNombre: safeStr(item?.nombre),
        tipo: 'entrada',
        cantidad: cant, motivo,
        stockAntes, stockDespues,
        fecha: serverTimestamp(),
        registradoPor: session.uid,
        registradoPorNombre: session.displayName,
      });

      ov.remove();
      showToast(`Entrada de ${cant} ${safeStr(item?.unidad,'')} registrada.`, 'success');
      await showInventario(db, session);
    } catch(e) {
      errEl.textContent = 'Error al registrar. Intenta de nuevo.';
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
    .map(d => ({ id: d.id, ...d.data() }))
    .filter(i => esItemValido(i) && safeNum(i.stock) > 0)
    .sort((a,b) => safeStr(a.nombre).localeCompare(safeStr(b.nombre)));

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
            ${items.map(i => `<option value="${i.id}"
              data-nombre="${safeStr(i.nombre)}"
              data-unidad="${safeStr(i.unidad)}"
              data-stock="${safeNum(i.stock)}"
            >${safeStr(i.nombre)} (${safeNum(i.stock)} ${safeStr(i.unidad)})</option>`).join('')}
          </select>
          <input id="fs-cant" type="number" min="1" placeholder="Cant."
            class="w-20 border border-gray-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 text-center shrink-0"/>
          <button id="fs-add"
            class="px-3 py-2.5 text-white rounded-lg text-sm font-bold shrink-0" style="background-color:#2196F3">
            +
          </button>
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
    if (sel.length === 0) {
      lista.innerHTML = '<p class="text-xs text-gray-400 text-center py-2">Sin materiales agregados</p>';
      return;
    }
    lista.innerHTML = sel.map((s, i) => `
      <div class="flex items-center justify-between bg-gray-50 rounded-lg px-3 py-2 gap-2">
        <div class="min-w-0">
          <p class="text-sm font-medium text-gray-900 truncate">${s.nombre}</p>
          <p class="text-xs text-gray-500">${s.cantidad} ${s.unidad}</p>
        </div>
        <button data-ri="${i}" class="text-gray-400 hover:text-red-500 shrink-0 text-lg leading-none">✕</button>
      </div>`).join('');
    lista.querySelectorAll('[data-ri]').forEach(b => {
      b.onclick = () => { sel.splice(parseInt(b.dataset.ri), 1); renderLista(); };
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
    if (cant <= 0) { errEl.textContent = 'Ingresa una cantidad válida.'; errEl.classList.remove('hidden'); return; }

    const disponible = safeNum(opt.dataset.stock);
    if (cant > disponible) {
      errEl.textContent = `Stock insuficiente. Disponible: ${disponible} ${opt.dataset.unidad}`;
      errEl.classList.remove('hidden');
      return;
    }

    const ex = sel.findIndex(s => s.itemId === selEl.value);
    if (ex >= 0) sel[ex].cantidad += cant;
    else sel.push({
      itemId: selEl.value,
      nombre: safeStr(opt.dataset.nombre),
      unidad: safeStr(opt.dataset.unidad),
      cantidad: cant,
    });

    selEl.value = ''; ov.querySelector('#fs-cant').value = '';
    renderLista();
  };

  ov.querySelector('#fs-submit').onclick = async () => {
    const tecSel  = ov.querySelector('#fs-tec');
    const tecUid  = tecSel.value;
    const tecNom  = safeStr(tecSel.options[tecSel.selectedIndex]?.dataset.nombre);
    const motivo  = ov.querySelector('#fs-motivo').value.trim();
    const errEl   = ov.querySelector('#fs-err');
    const btn     = ov.querySelector('#fs-submit');

    errEl.classList.add('hidden');
    if (!tecUid) { errEl.textContent = 'Selecciona el técnico.'; errEl.classList.remove('hidden'); return; }
    if (sel.length === 0) { errEl.textContent = 'Agrega al menos un material.'; errEl.classList.remove('hidden'); return; }

    btn.disabled = true; btn.textContent = 'Registrando...';
    try {
      const ref = await addDoc(collection(db, 'kardex/movimientos/salidas'), {
        tecnicoUid: tecUid, tecnicoNombre: tecNom, motivo,
        items: sel.map(s => ({ itemId: s.itemId, nombre: s.nombre, unidad: s.unidad, cantidad: s.cantidad })),
        fecha: serverTimestamp(),
        registradoPor: session.uid,
        registradoPorNombre: session.displayName,
      });

      for (const s of sel) {
        await updateDoc(doc(db, 'kardex/inventario/items', s.itemId), {
          stock: increment(-s.cantidad),
        });
      }

      ov.remove();
      showToast('Salida registrada correctamente.', 'success');
      showMemo({
        id: ref.id, tecnicoNombre: tecNom, motivo,
        items: sel, registradoPorNombre: session.displayName, fecha: new Date(),
      });
      await showDashboard(db, session);
    } catch(e) {
      errEl.textContent = 'Error al registrar. Intenta de nuevo.';
      errEl.classList.remove('hidden');
      btn.disabled = false; btn.textContent = 'Registrar salida';
      console.error(e);
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
      q = query(collection(db, 'kardex/movimientos/salidas'),
        where('tecnicoUid', '==', session.uid), orderBy('fecha', 'desc'));
    } else {
      q = query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'));
    }

    const snap   = await getDocs(q);
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
                  class="text-xs font-medium px-2 py-1 rounded-lg shrink-0"
                  style="color:#1B4F8A;background:#EFF6FF">
                  Memo
                </button>
              </div>
              ${s.motivo ? `<p class="text-xs text-gray-500 mt-1">${s.motivo}</p>` : ''}
              <div class="flex flex-wrap gap-1 mt-2">
                ${(s.items||[]).map(i =>
                  `<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">
                    ${safeNum(i.cantidad)} ${safeStr(i.unidad,'')} ${safeStr(i.nombre)}
                  </span>`
                ).join('')}
              </div>
            </div>`).join('') + '</div>'
        }
      </div>`;

    // Guardar datos de salidas en memoria para el memo
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
                <td class="px-3 py-2 text-gray-900">${safeStr(i.nombre)}</td>
                <td class="px-3 py-2 text-center font-semibold text-gray-900">${safeNum(i.cantidad)}</td>
                <td class="px-3 py-2 text-gray-500">${safeStr(i.unidad,'')}</td>
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
      <button id="fm-print" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">🖨️ Imprimir</button>
    </div>`);

  ov.querySelector('#fm-close').onclick = ov.querySelector('#fm-cancel').onclick = () => ov.remove();
  ov.querySelector('#fm-print').onclick = () => {
    const body = document.getElementById('memo-body').innerHTML;
    const w = window.open('', '_blank');
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
  ov.className = 'fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 overflow-y-auto';
  ov.innerHTML = `<div class="bg-white rounded-2xl shadow-xl w-full max-w-md my-4">${inner}</div>`;
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
