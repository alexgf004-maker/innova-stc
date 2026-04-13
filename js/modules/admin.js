/**
 * admin.js
 * Módulo de administración de usuarios.
 * Solo accesible con rol "admin".
 */

import { getFirestore, collection, getDocs, doc, setDoc,
         updateDoc, serverTimestamp, query, where }
  from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';
import { getAuth, createUserWithEmailAndPassword, updatePassword }
  from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js';
import { showToast, showModal, showInputModal,
         roleBadge, activeBadge, setButtonLoading }
  from '../ui.js';
import { generateSalt, hashPin, derivePassword } from '../crypto.js';
import { SEED } from '../firebase-config.js';

// ─────────────────────────────────────────
// INIT — llamado por router.js al cargar la vista
// ─────────────────────────────────────────
export async function initAdmin(session) {
  if (session.role !== 'admin') return;

  const db   = window.__firebase.db;
  const auth = window.__firebase.auth;

  await renderUserList(db);
  bindAdminEvents(db, auth, session);
}

// ─────────────────────────────────────────
// RENDERIZADO DE LISTA
// ─────────────────────────────────────────
async function renderUserList(db) {
  const container = document.getElementById('user-list-container');
  if (!container) return;

  container.innerHTML = loadingRow();

  try {
    const snap  = await getDocs(collection(db, 'users'));
    const users = snap.docs.map(d => d.data());

    if (users.length === 0) {
      container.innerHTML = emptyRow('No hay usuarios registrados.');
      return;
    }

    container.innerHTML = users.map(u => userRow(u)).join('');

  } catch (err) {
    container.innerHTML = emptyRow('Error al cargar usuarios.');
    console.error(err);
  }
}

function userRow(u) {
  return `
    <tr class="hover:bg-gray-50 transition-colors">
      <td class="px-4 py-3">
        <div class="font-medium text-gray-900 text-sm">${u.displayName}</div>
        <div class="text-xs text-gray-400 font-mono">${u.username}</div>
      </td>
      <td class="px-4 py-3 hidden sm:table-cell">${roleBadge(u.role)}</td>
      <td class="px-4 py-3">${activeBadge(u.active)}</td>
      <td class="px-4 py-3 text-right">
        <div class="flex items-center justify-end gap-2">
          <button data-action="reset-pin" data-uid="${u.uid}" data-name="${u.displayName}"
            title="Resetear PIN"
            class="p-1.5 text-gray-400 hover:text-accent transition-colors rounded">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="23 4 23 10 17 10"/>
              <path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/>
            </svg>
          </button>
          ${u.active ? `
          <button data-action="deactivate" data-uid="${u.uid}" data-name="${u.displayName}"
            title="Desactivar usuario"
            class="p-1.5 text-gray-400 hover:text-danger transition-colors rounded">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <circle cx="12" cy="12" r="10"/><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"/>
            </svg>
          </button>` : `
          <button data-action="activate" data-uid="${u.uid}" data-name="${u.displayName}"
            title="Activar usuario"
            class="p-1.5 text-gray-400 hover:text-success transition-colors rounded">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="20 6 9 17 4 12"/>
            </svg>
          </button>`}
        </div>
      </td>
    </tr>
  `;
}

function loadingRow() {
  return `<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400 text-sm">
    <div class="flex items-center justify-center gap-2">
      <svg class="animate-spin w-4 h-4" viewBox="0 0 24 24" fill="none">
        <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
        <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
      </svg>
      Cargando usuarios...
    </div>
  </td></tr>`;
}

function emptyRow(msg) {
  return `<tr><td colspan="4" class="px-4 py-8 text-center text-gray-400 text-sm">${msg}</td></tr>`;
}

// ─────────────────────────────────────────
// EVENTOS
// ─────────────────────────────────────────
function bindAdminEvents(db, auth, session) {
  // Botón "Nuevo usuario"
  const btnNew = document.getElementById('btn-new-user');
  if (btnNew) btnNew.addEventListener('click', () => showUserForm(db, auth, session));

  // Acciones en la tabla (delegación de eventos)
  const tableBody = document.getElementById('user-list-container');
  if (tableBody) {
    tableBody.addEventListener('click', async (e) => {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;

      const { action, uid, name } = btn.dataset;

      if (action === 'reset-pin')  await handleResetPin(db, uid, name);
      if (action === 'deactivate') await handleSetActive(db, uid, name, false);
      if (action === 'activate')   await handleSetActive(db, uid, name, true);

      await renderUserList(db);
    });
  }
}

// ─────────────────────────────────────────
// ACCIONES
// ─────────────────────────────────────────
async function handleResetPin(db, uid, name) {
  const newPin = await showInputModal({
    title:       `Resetear PIN — ${name}`,
    message:     'Ingresa el nuevo PIN (mínimo 4 dígitos numéricos).',
    placeholder: 'Nuevo PIN',
    inputType:   'password',
    inputMode:   'numeric',
  });

  if (!newPin) return;
  const cleaned = newPin.replace(/\D/g, '');
  if (cleaned.length < 4) { showToast('El PIN debe tener al menos 4 dígitos.', 'error'); return; }

  try {
    const salt    = generateSalt();
    const pinHash = await hashPin(salt, cleaned);
    await updateDoc(doc(db, 'users', uid), { pinHash, pinSalt: salt });
    showToast(`PIN de ${name} reseteado correctamente.`, 'success');
  } catch (err) {
    showToast('Error al resetear el PIN.', 'error');
    console.error(err);
  }
}

async function handleSetActive(db, uid, name, active) {
  const action = active ? 'activar' : 'desactivar';
  const confirmed = await showModal({
    title:        `${active ? 'Activar' : 'Desactivar'} usuario`,
    message:      `¿Confirmas que deseas ${action} a ${name}?`,
    confirmLabel: active ? 'Activar' : 'Desactivar',
    confirmType:  active ? 'primary' : 'danger',
  });

  if (!confirmed) return;

  try {
    await updateDoc(doc(db, 'users', uid), { active });
    showToast(`${name} ha sido ${active ? 'activado' : 'desactivado'}.`, active ? 'success' : 'info');
  } catch (err) {
    showToast('Error al actualizar el usuario.', 'error');
    console.error(err);
  }
}

// ─────────────────────────────────────────
// FORMULARIO DE NUEVO USUARIO
// ─────────────────────────────────────────
function showUserForm(db, auth, session) {
  const overlay = document.createElement('div');
  overlay.className = 'fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4 overflow-y-auto';
  overlay.innerHTML = `
    <div class="bg-white rounded-2xl shadow-xl w-full max-w-md my-4">
      <div class="flex items-center justify-between px-6 py-5 border-b border-gray-100">
        <h2 class="font-semibold text-gray-900">Nuevo usuario</h2>
        <button id="close-form" class="text-gray-400 hover:text-gray-700 transition-colors">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
          </svg>
        </button>
      </div>
      <div class="px-6 py-5 space-y-4">
        ${formField('displayName', 'Nombre completo', 'text', 'Ej. Bryan Francia')}
        ${formField('username', 'Usuario', 'text', 'Ej. bfrancia')}
        <div>
          <label class="block text-sm font-medium text-gray-700 mb-1.5">Rol</label>
          <select id="field-role"
            class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-accent">
            <option value="campo">Campo</option>
            <option value="coordinadora">Coordinadora</option>
            <option value="admin">Administrador</option>
          </select>
        </div>
        ${formField('pin', 'PIN inicial', 'password', 'Mínimo 4 dígitos', 'numeric')}
        <div id="form-error" class="hidden text-sm text-danger bg-red-50 border border-red-200 rounded-lg px-3.5 py-2.5"></div>
      </div>
      <div class="px-6 py-4 border-t border-gray-100 flex gap-3">
        <button id="cancel-form"
          class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50 transition-colors">
          Cancelar
        </button>
        <button id="submit-form"
          class="flex-1 bg-primary hover:bg-accent text-white font-medium rounded-lg py-2.5 text-sm transition-colors">
          Crear usuario
        </button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  overlay.querySelector('#close-form').addEventListener('click',  () => overlay.remove());
  overlay.querySelector('#cancel-form').addEventListener('click', () => overlay.remove());

  // Solo dígitos en PIN
  overlay.querySelector('#field-pin').addEventListener('input', (e) => {
    e.target.value = e.target.value.replace(/\D/g, '').slice(0, 8);
  });

  overlay.querySelector('#submit-form').addEventListener('click', async () => {
    await handleCreateUser(db, auth, overlay, session);
  });
}

function formField(id, label, type, placeholder, inputMode = '') {
  return `
    <div>
      <label class="block text-sm font-medium text-gray-700 mb-1.5">${label}</label>
      <input id="field-${id}" type="${type}" placeholder="${placeholder}"
        ${inputMode ? `inputmode="${inputMode}"` : ''}
        class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-accent" />
    </div>
  `;
}

async function handleCreateUser(db, auth, overlay, session) {
  const displayName = overlay.querySelector('#field-displayName').value.trim();
  const username    = overlay.querySelector('#field-username').value.trim().toLowerCase();
  const role        = overlay.querySelector('#field-role').value;
  const pin         = overlay.querySelector('#field-pin').value.replace(/\D/g, '');
  const errorEl     = overlay.querySelector('#form-error');
  const submitBtn   = overlay.querySelector('#submit-form');

  errorEl.classList.add('hidden');

  // Validaciones básicas
  if (!displayName || !username || !pin) {
    errorEl.textContent = 'Todos los campos son obligatorios.';
    errorEl.classList.remove('hidden');
    return;
  }
  if (pin.length < 4) {
    errorEl.textContent = 'El PIN debe tener al menos 4 dígitos.';
    errorEl.classList.remove('hidden');
    return;
  }
  if (!/^[a-z0-9._]+$/.test(username)) {
    errorEl.textContent = 'El usuario solo puede contener letras, números, puntos y guiones bajos.';
    errorEl.classList.remove('hidden');
    return;
  }

  // Verificar que el username no exista ya
  try {
    const q    = query(collection(db, 'users'), where('username', '==', username));
    const snap = await getDocs(q);
    if (!snap.empty) {
      errorEl.textContent = 'Ese nombre de usuario ya está en uso.';
      errorEl.classList.remove('hidden');
      return;
    }
  } catch (err) {
    errorEl.textContent = 'Error al verificar el usuario.';
    errorEl.classList.remove('hidden');
    return;
  }

  submitBtn.disabled = true;
  submitBtn.innerHTML = `<svg class="animate-spin w-4 h-4 mx-auto" viewBox="0 0 24 24" fill="none">
    <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
    <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
  </svg>`;

  try {
    const internalEmail = `${username}@innova-stc.internal`;

    // Crear segunda instancia de Firebase para no perder la sesión del admin
    const { initializeApp, getApp } = await import('https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js');
    const { getAuth, createUserWithEmailAndPassword, updatePassword } = await import('https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js');

    let secondApp;
    try { secondApp = getApp('secondary'); } catch { 
      secondApp = initializeApp(window.__firebase.app.options, 'secondary');
    }
    const secondAuth = getAuth(secondApp);

    const cred = await createUserWithEmailAndPassword(secondAuth, internalEmail, 'TEMP_PASS_placeholder');
    const uid  = cred.user.uid;

    const realPass = await derivePassword(uid, SEED);
    await updatePassword(cred.user, realPass);

    const salt    = generateSalt();
    const pinHash = await hashPin(salt, pin);

    await setDoc(doc(db, 'users', uid), {
      uid, username, displayName, role, internalEmail,
      pinHash, pinSalt: salt, active: true,
      createdAt: serverTimestamp(), createdBy: session.uid,
    });

    overlay.remove();
    showToast(`Usuario ${displayName} creado correctamente.`, 'success');
    await renderUserList(db);

  } catch (err) {
    submitBtn.disabled = false;
    submitBtn.textContent = 'Crear usuario';
    errorEl.textContent = err.message?.includes('email-already-in-use')
      ? 'Ese nombre de usuario ya está registrado en el sistema.'
      : 'Error al crear el usuario. Intenta de nuevo.';
    errorEl.classList.remove('hidden');
    console.error(err);
  }
}
