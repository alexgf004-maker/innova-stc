/**
 * kardex.js — Fase 1
 * Control de inventario y salidas de materiales.
 */

import {
  collection, doc, getDocs, addDoc, updateDoc,
  query, orderBy, where, serverTimestamp, increment
} from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';

import { showToast } from '../ui.js';

export async function initKardex(session) {
  const container = document.getElementById('kardex-root');
  if (!container) return;
  const db = window.__firebase.db;
  renderShell(container, session);
  await showDashboard(db, session);
  bindNav(db, session);
}

function renderShell(container, session) {
  const canEdit = ['admin','coordinadora'].includes(session.role);
  container.innerHTML = `
    <div class="space-y-4">
      <div class="flex items-center justify-between">
        <div>
          <h1 class="text-xl font-semibold text-gray-900">Kardex</h1>
          <p class="text-sm text-gray-500 mt-0.5">Control de materiales</p>
        </div>
        ${canEdit ? `<button id="btn-nueva-salida" class="inline-flex items-center gap-2 text-white text-sm font-medium px-4 py-2.5 rounded-lg" style="background-color:#1B4F8A">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
          Nueva salida
        </button>` : ''}
      </div>
      <div class="flex gap-1 bg-gray-100 rounded-xl p-1">
        <button data-tab="dashboard" class="ktab flex-1 py-2 text-sm font-medium rounded-lg transition-colors">Inicio</button>
        <button data-tab="inventario" class="ktab flex-1 py-2 text-sm font-medium rounded-lg transition-colors">Inventario</button>
        <button data-tab="historial" class="ktab flex-1 py-2 text-sm font-medium rounded-lg transition-colors">Historial</button>
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
      if (t === 'dashboard') await showDashboard(db, session);
      if (t === 'inventario') await showInventario(db, session);
      if (t === 'historial') await showHistorial(db, session);
    });
  });
  document.getElementById('btn-nueva-salida')?.addEventListener('click', () => showFormSalida(db, session));
}

// ── DASHBOARD ──────────────────────────────────────────────
async function showDashboard(db, session) {
  setTab('dashboard');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  try {
    const snap = await getDocs(collection(db, 'kardex/inventario/items'));
    const items = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    const sinStock = items.filter(i => i.stock <= 0).length;
    const stockBajo = items.filter(i => i.stock > 0 && i.stock <= (i.stockMinimo || 5)).length;

    const qS = query(collection(db, 'kardex/movimientos/salidas'), orderBy('fecha', 'desc'));
    const snapS = await getDocs(qS);
    const ultimas = snapS.docs.slice(0, 5).map(d => ({ id: d.id, ...d.data() }));

    content.innerHTML = `
      <div class="space-y-4">
        <div class="grid grid-cols-3 gap-3">
          ${stat('Materiales', items.length, '#2196F3')}
          ${stat('Stock bajo', stockBajo, '#E65100')}
          ${stat('Sin stock', sinStock, '#C62828')}
        </div>
        <div class="bg-white rounded-xl border border-gray-200 overflow-hidden">
          <div class="px-4 py-3 border-b border-gray-100 flex items-center justify-between">
            <p class="font-medium text-sm text-gray-900">Últimas salidas</p>
            <button data-tab="historial" class="ktab text-xs font-medium" style="color:#2196F3">Ver todo</button>
          </div>
          ${ultimas.length === 0
            ? '<div class="px-4 py-6 text-center text-sm text-gray-400">Sin movimientos aún</div>'
            : '<div class="divide-y divide-gray-100">' + ultimas.map(s => `
              <div class="px-4 py-3">
                <div class="flex items-start justify-between gap-2">
                  <div class="min-w-0">
                    <p class="text-sm font-medium text-gray-900 truncate">${s.tecnicoNombre}</p>
                    <p class="text-xs text-gray-400 mt-0.5">${fmtDate(s.fecha)}</p>
                  </div>
                  <p class="text-xs text-gray-500 shrink-0">${s.items?.length || 0} ítem(s)</p>
                </div>
                ${s.motivo ? `<p class="text-xs text-gray-400 mt-1 truncate">${s.motivo}</p>` : ''}
              </div>`).join('') + '</div>'
          }
        </div>
        ${stockBajo > 0 || sinStock > 0 ? `
        <div class="bg-white rounded-xl border border-orange-200 overflow-hidden">
          <div class="px-4 py-3 border-b border-orange-100">
            <p class="font-medium text-sm text-orange-800">⚠️ Atención al inventario</p>
          </div>
          <div class="divide-y divide-gray-100">
            ${items.filter(i => i.stock <= (i.stockMinimo || 5)).map(i => `
              <div class="px-4 py-3 flex items-center justify-between">
                <p class="text-sm text-gray-900">${i.nombre}</p>
                <span class="text-xs font-medium px-2 py-1 rounded-full ${i.stock <= 0 ? 'bg-red-100 text-red-700' : 'bg-orange-100 text-orange-700'}">
                  ${i.stock <= 0 ? 'Sin stock' : i.stock + ' ' + i.unidad}
                </span>
              </div>`).join('')}
          </div>
        </div>` : ''}
      </div>`;

    document.querySelectorAll('.ktab').forEach(b => {
      b.addEventListener('click', async () => {
        const t = b.dataset.tab;
        if (t === 'dashboard') await showDashboard(db, session);
        if (t === 'inventario') await showInventario(db, session);
        if (t === 'historial') await showHistorial(db, session);
      });
    });
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

function stat(label, val, color) {
  return `<div class="bg-white rounded-xl border border-gray-200 p-4 text-center">
    <p class="text-2xl font-bold" style="color:${color}">${val}</p>
    <p class="text-xs text-gray-500 mt-1">${label}</p>
  </div>`;
}

// ── INVENTARIO ─────────────────────────────────────────────
async function showInventario(db, session) {
  setTab('inventario');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  const canEdit = ['admin','coordinadora'].includes(session.role);
  try {
    const snap = await getDocs(collection(db, 'kardex/inventario/items'));
    const items = snap.docs.map(d => ({ id: d.id, ...d.data() })).sort((a,b) => a.nombre.localeCompare(b.nombre));

    content.innerHTML = `
      <div class="space-y-3">
        ${canEdit ? `<div class="flex justify-end">
          <button id="btn-nuevo-item" class="inline-flex items-center gap-2 text-white text-sm font-medium px-3 py-2 rounded-lg" style="background-color:#1B4F8A">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
            Agregar material
          </button>
        </div>` : ''}
        <div class="bg-white rounded-xl border border-gray-200 overflow-hidden">
          ${items.length === 0
            ? `<div class="py-12 text-center text-sm text-gray-400">No hay materiales registrados.${canEdit ? '<br>Agrega el primero.' : ''}</div>`
            : `<div class="overflow-x-auto"><table class="w-full text-sm">
                <thead class="bg-gray-50 border-b border-gray-200">
                  <tr>
                    <th class="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Material</th>
                    <th class="text-center px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wide">Stock</th>
                    ${canEdit ? '<th class="px-4 py-3"></th>' : ''}
                  </tr>
                </thead>
                <tbody class="divide-y divide-gray-100">
                  ${items.map(item => {
                    const sc = item.stock <= 0
                      ? 'color:#C62828;background:#FEE2E2'
                      : item.stock <= (item.stockMinimo||5)
                      ? 'color:#E65100;background:#FEF3C7'
                      : 'color:#166534;background:#DCFCE7';
                    return `<tr class="hover:bg-gray-50">
                      <td class="px-4 py-3"><p class="font-medium text-gray-900">${item.nombre}</p><p class="text-xs text-gray-400">${item.unidad}</p></td>
                      <td class="px-4 py-3 text-center"><span class="text-xs font-semibold px-2 py-1 rounded-full" style="${sc}">${item.stock} ${item.unidad}</span></td>
                      ${canEdit ? `<td class="px-4 py-3"><div class="flex items-center justify-end gap-2">
                        <button data-ajuste="${item.id}" title="Ajustar stock" class="p-1.5 text-gray-400 hover:text-blue-500 rounded">
                          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>
                        </button>
                        <button data-edit="${item.id}" title="Editar" class="p-1.5 text-gray-400 hover:text-blue-500 rounded">
                          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                      </div></td>` : ''}
                    </tr>`;
                  }).join('')}
                </tbody>
              </table></div>`
          }
        </div>
      </div>`;

    if (canEdit) {
      document.getElementById('btn-nuevo-item')?.addEventListener('click', () => showFormItem(db, session, null));
      content.querySelectorAll('[data-edit]').forEach(b => {
        b.addEventListener('click', () => showFormItem(db, session, items.find(i => i.id === b.dataset.edit)));
      });
      content.querySelectorAll('[data-ajuste]').forEach(b => {
        b.addEventListener('click', () => showFormAjuste(db, session, items.find(i => i.id === b.dataset.ajuste)));
      });
    }
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ── FORM ITEM ──────────────────────────────────────────────
function showFormItem(db, session, item) {
  const esNuevo = !item;
  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">${esNuevo ? 'Nuevo material' : 'Editar material'}</h2>
      <button id="fi-close" class="text-gray-400 hover:text-gray-700"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
    </div>
    <div class="px-5 py-4 space-y-4">
      <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Nombre</label>
        <input id="fi-nombre" type="text" value="${item?.nombre||''}" placeholder="Ej. Medidor monofásico" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Unidad</label>
          <input id="fi-unidad" type="text" value="${item?.unidad||''}" placeholder="unidad, m, kg" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>
        <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Stock mínimo</label>
          <input id="fi-min" type="number" min="0" value="${item?.stockMinimo||5}" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>
      </div>
      ${esNuevo ? `<div><label class="block text-sm font-medium text-gray-700 mb-1.5">Stock inicial</label>
        <input id="fi-stock" type="number" min="0" value="0" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>` : ''}
      <div id="fi-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fi-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cancelar</button>
      <button id="fi-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">${esNuevo?'Agregar':'Guardar'}</button>
    </div>`);

  ov.querySelector('#fi-close').onclick = ov.querySelector('#fi-cancel').onclick = () => ov.remove();
  ov.querySelector('#fi-submit').onclick = async () => {
    const nombre = ov.querySelector('#fi-nombre').value.trim();
    const unidad = ov.querySelector('#fi-unidad').value.trim();
    const minimo = parseInt(ov.querySelector('#fi-min').value)||0;
    const errEl  = ov.querySelector('#fi-err');
    const btn    = ov.querySelector('#fi-submit');
    errEl.classList.add('hidden');
    if (!nombre||!unidad) { errEl.textContent='Nombre y unidad obligatorios.'; errEl.classList.remove('hidden'); return; }
    btn.disabled=true; btn.textContent='Guardando...';
    try {
      if (esNuevo) {
        const stock = parseInt(ov.querySelector('#fi-stock').value)||0;
        await addDoc(collection(db,'kardex/inventario/items'), { nombre, unidad, stock, stockMinimo:minimo, creadoEn:serverTimestamp(), creadoPor:session.uid });
      } else {
        await updateDoc(doc(db,'kardex/inventario/items',item.id), { nombre, unidad, stockMinimo:minimo });
      }
      ov.remove(); showToast(`Material ${esNuevo?'agregado':'actualizado'}.`,'success');
      await showInventario(db, session);
    } catch(e) { errEl.textContent='Error al guardar.'; errEl.classList.remove('hidden'); btn.disabled=false; btn.textContent=esNuevo?'Agregar':'Guardar'; }
  };
}

// ── FORM AJUSTE ────────────────────────────────────────────
function showFormAjuste(db, session, item) {
  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Ajustar stock</h2>
      <button id="fa-close" class="text-gray-400 hover:text-gray-700"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
    </div>
    <div class="px-5 py-4 space-y-4">
      <div class="bg-gray-50 rounded-xl p-3">
        <p class="text-sm font-medium text-gray-900">${item.nombre}</p>
        <p class="text-xs text-gray-500 mt-0.5">Stock actual: <strong>${item.stock} ${item.unidad}</strong></p>
      </div>
      <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Tipo</label>
        <select id="fa-tipo" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400">
          <option value="entrada">Entrada (sumar al stock)</option>
          <option value="ajuste">Corrección (establecer cantidad exacta)</option>
        </select></div>
      <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Cantidad</label>
        <input id="fa-cant" type="number" min="0" placeholder="0" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>
      <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Motivo (opcional)</label>
        <input id="fa-motivo" type="text" placeholder="Ej. Compra de materiales" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>
      <div id="fa-err" class="hidden text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fa-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cancelar</button>
      <button id="fa-submit" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">Aplicar</button>
    </div>`);

  ov.querySelector('#fa-close').onclick = ov.querySelector('#fa-cancel').onclick = () => ov.remove();
  ov.querySelector('#fa-submit').onclick = async () => {
    const tipo   = ov.querySelector('#fa-tipo').value;
    const cant   = parseInt(ov.querySelector('#fa-cant').value);
    const motivo = ov.querySelector('#fa-motivo').value.trim();
    const errEl  = ov.querySelector('#fa-err');
    const btn    = ov.querySelector('#fa-submit');
    errEl.classList.add('hidden');
    if (isNaN(cant)||cant<0) { errEl.textContent='Cantidad inválida.'; errEl.classList.remove('hidden'); return; }
    btn.disabled=true; btn.textContent='Aplicando...';
    try {
      const delta = tipo==='entrada' ? cant : (cant - item.stock);
      await updateDoc(doc(db,'kardex/inventario/items',item.id), { stock: increment(delta) });
      await addDoc(collection(db,'kardex/movimientos/ajustes'), {
        itemId:item.id, itemNombre:item.nombre, tipo, cantidad:cant, motivo,
        stockAntes:item.stock, stockDespues: tipo==='entrada' ? item.stock+cant : cant,
        fecha:serverTimestamp(), registradoPor:session.uid, registradoPorNombre:session.displayName
      });
      ov.remove(); showToast('Stock actualizado.','success');
      await showInventario(db, session);
    } catch(e) { errEl.textContent='Error al ajustar.'; errEl.classList.remove('hidden'); btn.disabled=false; btn.textContent='Aplicar'; }
  };
}

// ── FORM SALIDA ────────────────────────────────────────────
async function showFormSalida(db, session) {
  const [snapI, snapU] = await Promise.all([
    getDocs(collection(db,'kardex/inventario/items')),
    getDocs(query(collection(db,'users'), where('active','==',true)))
  ]);
  const items    = snapI.docs.map(d=>({id:d.id,...d.data()})).filter(i=>i.stock>0).sort((a,b)=>a.nombre.localeCompare(b.nombre));
  const tecnicos = snapU.docs.map(d=>d.data()).filter(u=>u.role==='campo').sort((a,b)=>a.displayName.localeCompare(b.displayName));
  let sel = [];

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Nueva salida de materiales</h2>
      <button id="fs-close" class="text-gray-400 hover:text-gray-700"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
    </div>
    <div class="px-5 py-4 space-y-4">
      <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Técnico que recibe</label>
        <select id="fs-tec" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400">
          <option value="">Seleccionar técnico...</option>
          ${tecnicos.map(t=>`<option value="${t.uid}" data-nombre="${t.displayName}">${t.displayName}</option>`).join('')}
        </select></div>
      <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Motivo / Trabajo</label>
        <input id="fs-motivo" type="text" placeholder="Ej. Cambio de medidor, Servicio nuevo" class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"/></div>
      <div><label class="block text-sm font-medium text-gray-700 mb-1.5">Materiales</label>
        <div class="flex gap-2">
          <select id="fs-item" class="flex-1 border border-gray-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400">
            <option value="">Seleccionar material...</option>
            ${items.map(i=>`<option value="${i.id}" data-nombre="${i.nombre}" data-unidad="${i.unidad}" data-stock="${i.stock}">${i.nombre} (${i.stock} ${i.unidad})</option>`).join('')}
          </select>
          <input id="fs-cant" type="number" min="1" placeholder="Cant." class="w-20 border border-gray-300 rounded-lg px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 text-center"/>
          <button id="fs-add" class="px-3 py-2.5 text-white rounded-lg text-sm font-bold shrink-0" style="background-color:#2196F3">+</button>
        </div></div>
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
      : sel.map((s,i)=>`
        <div class="flex items-center justify-between bg-gray-50 rounded-lg px-3 py-2 gap-2">
          <div class="min-w-0"><p class="text-sm font-medium text-gray-900 truncate">${s.nombre}</p><p class="text-xs text-gray-500">${s.cantidad} ${s.unidad}</p></div>
          <button data-ri="${i}" class="text-gray-400 hover:text-red-500 shrink-0">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>`).join('');
    lista.querySelectorAll('[data-ri]').forEach(b => {
      b.onclick = () => { sel.splice(parseInt(b.dataset.ri),1); renderLista(); };
    });
  }
  renderLista();

  ov.querySelector('#fs-close').onclick = ov.querySelector('#fs-cancel').onclick = () => ov.remove();
  ov.querySelector('#fs-add').onclick = () => {
    const sel2 = ov.querySelector('#fs-item');
    const opt  = sel2.options[sel2.selectedIndex];
    const cant = parseInt(ov.querySelector('#fs-cant').value);
    const errEl = ov.querySelector('#fs-err');
    errEl.classList.add('hidden');
    if (!sel2.value) { errEl.textContent='Selecciona un material.'; errEl.classList.remove('hidden'); return; }
    if (!cant||cant<=0) { errEl.textContent='Cantidad inválida.'; errEl.classList.remove('hidden'); return; }
    if (cant > parseInt(opt.dataset.stock)) { errEl.textContent=`Stock insuficiente. Disponible: ${opt.dataset.stock} ${opt.dataset.unidad}`; errEl.classList.remove('hidden'); return; }
    const ex = sel.findIndex(s=>s.itemId===sel2.value);
    if (ex>=0) sel[ex].cantidad+=cant;
    else sel.push({ itemId:sel2.value, nombre:opt.dataset.nombre, unidad:opt.dataset.unidad, cantidad:cant, stockDisponible:parseInt(opt.dataset.stock) });
    sel2.value=''; ov.querySelector('#fs-cant').value='';
    renderLista();
  };

  ov.querySelector('#fs-submit').onclick = async () => {
    const tecSel  = ov.querySelector('#fs-tec');
    const tecUid  = tecSel.value;
    const tecNom  = tecSel.options[tecSel.selectedIndex]?.dataset.nombre||'';
    const motivo  = ov.querySelector('#fs-motivo').value.trim();
    const errEl   = ov.querySelector('#fs-err');
    const btn     = ov.querySelector('#fs-submit');
    errEl.classList.add('hidden');
    if (!tecUid) { errEl.textContent='Selecciona el técnico.'; errEl.classList.remove('hidden'); return; }
    if (sel.length===0) { errEl.textContent='Agrega al menos un material.'; errEl.classList.remove('hidden'); return; }
    btn.disabled=true; btn.textContent='Registrando...';
    try {
      const ref = await addDoc(collection(db,'kardex/movimientos/salidas'), {
        tecnicoUid:tecUid, tecnicoNombre:tecNom, motivo,
        items: sel.map(s=>({ itemId:s.itemId, nombre:s.nombre, unidad:s.unidad, cantidad:s.cantidad })),
        fecha: serverTimestamp(), registradoPor:session.uid, registradoPorNombre:session.displayName
      });
      for (const s of sel) {
        await updateDoc(doc(db,'kardex/inventario/items',s.itemId), { stock: increment(-s.cantidad) });
      }
      ov.remove();
      showToast('Salida registrada.','success');
      showMemo({ id:ref.id, tecnicoNombre:tecNom, motivo, items:sel, registradoPorNombre:session.displayName, fecha:new Date() });
      await showDashboard(db, session);
    } catch(e) { errEl.textContent='Error al registrar.'; errEl.classList.remove('hidden'); btn.disabled=false; btn.textContent='Registrar salida'; console.error(e); }
  };
}

// ── HISTORIAL ──────────────────────────────────────────────
async function showHistorial(db, session) {
  setTab('historial');
  const content = document.getElementById('kardex-content');
  content.innerHTML = loading();
  try {
    let q;
    if (session.role==='campo') {
      q = query(collection(db,'kardex/movimientos/salidas'), where('tecnicoUid','==',session.uid), orderBy('fecha','desc'));
    } else {
      q = query(collection(db,'kardex/movimientos/salidas'), orderBy('fecha','desc'));
    }
    const snap = await getDocs(q);
    const salidas = snap.docs.map(d=>({id:d.id,...d.data()}));

    content.innerHTML = `
      <div class="bg-white rounded-xl border border-gray-200 overflow-hidden">
        <div class="px-4 py-3 border-b border-gray-100">
          <p class="font-medium text-sm text-gray-900">Historial de salidas</p>
          <p class="text-xs text-gray-400 mt-0.5">${salidas.length} registro(s)</p>
        </div>
        ${salidas.length===0
          ? '<div class="py-12 text-center text-sm text-gray-400">Sin movimientos registrados</div>'
          : '<div class="divide-y divide-gray-100">' + salidas.map(s=>`
            <div class="px-4 py-3">
              <div class="flex items-start justify-between gap-2">
                <div class="min-w-0">
                  <p class="text-sm font-semibold text-gray-900">${s.tecnicoNombre}</p>
                  <p class="text-xs text-gray-400">${fmtDate(s.fecha)} · ${s.registradoPorNombre}</p>
                </div>
                <button data-smemo='${JSON.stringify({id:s.id,tecnicoNombre:s.tecnicoNombre,motivo:s.motivo,items:s.items,registradoPorNombre:s.registradoPorNombre,fecha:null})}'
                  class="text-xs font-medium px-2 py-1 rounded-lg shrink-0" style="color:#1B4F8A;background:#EFF6FF">Memo</button>
              </div>
              ${s.motivo?`<p class="text-xs text-gray-500 mt-1">${s.motivo}</p>`:''}
              <div class="flex flex-wrap gap-1 mt-2">
                ${(s.items||[]).map(i=>`<span class="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">${i.cantidad} ${i.unidad} ${i.nombre}</span>`).join('')}
              </div>
            </div>`).join('') + '</div>'
        }
      </div>`;

    content.querySelectorAll('[data-smemo]').forEach(b => {
      b.onclick = () => {
        const d = JSON.parse(b.dataset.smemo);
        d.fecha = new Date();
        showMemo(d);
      };
    });
  } catch(e) { content.innerHTML = errHtml(); console.error(e); }
}

// ── MEMO ───────────────────────────────────────────────────
function showMemo(s) {
  const fecha = s.fecha instanceof Date ? s.fecha : new Date();
  const fechaStr = fecha.toLocaleDateString('es-SV',{year:'numeric',month:'long',day:'numeric'});

  const ov = mkOverlay(`
    <div class="flex items-center justify-between px-5 py-4 border-b border-gray-100">
      <h2 class="font-semibold text-gray-900">Memo de salida</h2>
      <button id="fm-close" class="text-gray-400 hover:text-gray-700"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
    </div>
    <div id="memo-body" class="px-6 py-5 space-y-4">
      <div class="text-center border-b border-gray-200 pb-4">
        <p class="font-bold text-lg" style="color:#1B4F8A">INNOVA STC</p>
        <p class="text-sm text-gray-500">Servicios Técnicos y Comerciales</p>
        <p class="text-xs text-gray-400 mt-1">MEMO DE SALIDA DE MATERIALES</p>
      </div>
      <div class="grid grid-cols-2 gap-3 text-sm">
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">Fecha</p><p class="font-medium text-gray-900 mt-0.5">${fechaStr}</p></div>
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">N° Ref.</p><p class="font-medium text-gray-900 font-mono mt-0.5">${s.id.slice(-6).toUpperCase()}</p></div>
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">Técnico</p><p class="font-medium text-gray-900 mt-0.5">${s.tecnicoNombre}</p></div>
        <div><p class="text-xs text-gray-400 uppercase tracking-wide">Entregado por</p><p class="font-medium text-gray-900 mt-0.5">${s.registradoPorNombre}</p></div>
      </div>
      ${s.motivo?`<div><p class="text-xs text-gray-400 uppercase tracking-wide">Motivo / Trabajo</p><p class="text-sm font-medium text-gray-900 mt-0.5">${s.motivo}</p></div>`:''}
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
            ${(s.items||[]).map(i=>`<tr><td class="px-3 py-2 text-gray-900">${i.nombre}</td><td class="px-3 py-2 text-center font-semibold">${i.cantidad}</td><td class="px-3 py-2 text-gray-500">${i.unidad}</td></tr>`).join('')}
          </tbody>
        </table>
      </div>
      <div class="grid grid-cols-2 gap-6 pt-4">
        <div class="text-center"><div class="border-b border-gray-400 mb-1 h-10"></div><p class="text-xs text-gray-500">Entregado por</p><p class="text-xs font-medium text-gray-700">${s.registradoPorNombre}</p></div>
        <div class="text-center"><div class="border-b border-gray-400 mb-1 h-10"></div><p class="text-xs text-gray-500">Recibido por</p><p class="text-xs font-medium text-gray-700">${s.tecnicoNombre}</p></div>
      </div>
    </div>
    <div class="px-5 py-4 border-t border-gray-100 flex gap-3">
      <button id="fm-cancel" class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50">Cerrar</button>
      <button id="fm-print" class="flex-1 text-white font-medium rounded-lg py-2.5 text-sm" style="background-color:#1B4F8A">🖨️ Imprimir</button>
    </div>`);

  ov.querySelector('#fm-close').onclick = ov.querySelector('#fm-cancel').onclick = () => ov.remove();
  ov.querySelector('#fm-print').onclick = () => {
    const body = document.getElementById('memo-body').innerHTML;
    const w = window.open('','_blank');
    w.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Memo</title>
      <style>body{font-family:Arial,sans-serif;font-size:13px;color:#111;margin:24px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:6px 10px;text-align:left}th{background:#f5f5f5;font-size:11px}@media print{body{margin:10mm}}</style>
      </head><body>${body}</body></html>`);
    w.document.close(); w.print();
  };
}

// ── HELPERS ────────────────────────────────────────────────
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
  const d = ts.toDate ? ts.toDate() : new Date(ts);
  return d.toLocaleDateString('es-SV',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
}
