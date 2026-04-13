/**
 * ui.js
 * Helpers de interfaz reutilizables en toda la app.
 * Toast, Modal, Loading, Badges, helpers de formulario.
 */

export function initUI() {
  // El contenedor de toasts ya existe en index.html
}

// ─────────────────────────────────────────
// TOAST
// ─────────────────────────────────────────
/**
 * @param {string} message
 * @param {'success'|'error'|'info'|'warning'} type
 * @param {number} duration ms
 */
export function showToast(message, type = 'info', duration = 3500) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const colors = {
    success: 'bg-green-600',
    error:   'bg-red-600',
    warning: 'bg-orange-500',
    info:    'bg-primary',
  };

  const icons = {
    success: `<path d="M20 6L9 17l-5-5"/>`,
    error:   `<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>`,
    warning: `<path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>`,
    info:    `<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>`,
  };

  const toast = document.createElement('div');
  toast.className = `pointer-events-auto flex items-center gap-3 ${colors[type]} text-white text-sm font-medium px-4 py-3 rounded-xl shadow-lg max-w-sm translate-x-0 transition-all duration-300`;
  toast.innerHTML = `
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="shrink-0">
      ${icons[type]}
    </svg>
    <span>${message}</span>
  `;

  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(1rem)';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// ─────────────────────────────────────────
// MODAL
// ─────────────────────────────────────────
/**
 * Muestra un modal de confirmación.
 * @param {object} options
 * @param {string} options.title
 * @param {string} options.message
 * @param {string} options.confirmLabel
 * @param {'danger'|'primary'} options.confirmType
 * @returns {Promise<boolean>}
 */
export function showModal({ title, message, confirmLabel = 'Confirmar', confirmType = 'primary' }) {
  return new Promise((resolve) => {
    const btnColor = confirmType === 'danger'
      ? 'bg-red-600 hover:bg-red-700 text-white'
      : 'bg-primary hover:bg-accent text-white';

    const overlay = document.createElement('div');
    overlay.className = 'fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4';
    overlay.innerHTML = `
      <div class="bg-white rounded-2xl shadow-xl w-full max-w-sm p-6 space-y-4">
        <h3 class="font-semibold text-gray-900 text-base">${title}</h3>
        <p class="text-sm text-gray-600">${message}</p>
        <div class="flex gap-3 pt-2">
          <button id="modal-cancel"
            class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50 transition-colors">
            Cancelar
          </button>
          <button id="modal-confirm"
            class="flex-1 ${btnColor} font-medium rounded-lg py-2.5 text-sm transition-colors">
            ${confirmLabel}
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);

    overlay.querySelector('#modal-confirm').addEventListener('click', () => {
      overlay.remove();
      resolve(true);
    });
    overlay.querySelector('#modal-cancel').addEventListener('click', () => {
      overlay.remove();
      resolve(false);
    });
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) { overlay.remove(); resolve(false); }
    });
  });
}

/**
 * Modal de input (para resetear PIN, etc.)
 */
export function showInputModal({ title, message, placeholder = '', inputType = 'text', inputMode = 'text' }) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4';
    overlay.innerHTML = `
      <div class="bg-white rounded-2xl shadow-xl w-full max-w-sm p-6 space-y-4">
        <h3 class="font-semibold text-gray-900 text-base">${title}</h3>
        <p class="text-sm text-gray-600">${message}</p>
        <input
          id="modal-input"
          type="${inputType}"
          inputmode="${inputMode}"
          placeholder="${placeholder}"
          class="w-full border border-gray-300 rounded-lg px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-accent"
        />
        <div id="modal-error" class="hidden text-sm text-danger bg-red-50 border border-red-200 rounded-lg px-3 py-2"></div>
        <div class="flex gap-3 pt-2">
          <button id="modal-cancel"
            class="flex-1 border border-gray-300 text-gray-700 font-medium rounded-lg py-2.5 text-sm hover:bg-gray-50 transition-colors">
            Cancelar
          </button>
          <button id="modal-confirm"
            class="flex-1 bg-primary hover:bg-accent text-white font-medium rounded-lg py-2.5 text-sm transition-colors">
            Confirmar
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);
    const input     = overlay.querySelector('#modal-input');
    const errorEl   = overlay.querySelector('#modal-error');
    input.focus();

    overlay.querySelector('#modal-confirm').addEventListener('click', () => {
      const val = input.value.trim();
      if (!val) {
        errorEl.textContent = 'Este campo es obligatorio.';
        errorEl.classList.remove('hidden');
        return;
      }
      overlay.remove();
      resolve(val);
    });
    overlay.querySelector('#modal-cancel').addEventListener('click', () => {
      overlay.remove();
      resolve(null);
    });
  });
}

// ─────────────────────────────────────────
// BADGES
// ─────────────────────────────────────────
const ROLE_BADGE = {
  admin:        'badge-purple',
  coordinadora: 'badge-blue',
  campo:        'badge-green',
};

const STATUS_BADGE = {
  true:  'badge-green',
  false: 'badge-gray',
};

export function roleBadge(role) {
  const cls = ROLE_BADGE[role] || 'badge-gray';
  const labels = { admin: 'Admin', coordinadora: 'Coordinadora', campo: 'Campo' };
  return `<span class="badge ${cls}">${labels[role] || role}</span>`;
}

export function activeBadge(active) {
  const cls   = STATUS_BADGE[String(active)] || 'badge-gray';
  const label = active ? 'Activo' : 'Inactivo';
  return `<span class="badge ${cls}">${label}</span>`;
}

// ─────────────────────────────────────────
// FORM HELPERS
// ─────────────────────────────────────────
export function setFieldError(fieldId, message) {
  const field = document.getElementById(fieldId);
  const error = document.getElementById(fieldId + '-error');
  if (field) field.classList.add('border-red-400', 'focus:ring-red-400');
  if (error) { error.textContent = message; error.classList.remove('hidden'); }
}

export function clearFieldError(fieldId) {
  const field = document.getElementById(fieldId);
  const error = document.getElementById(fieldId + '-error');
  if (field) field.classList.remove('border-red-400', 'focus:ring-red-400');
  if (error) error.classList.add('hidden');
}

export function setButtonLoading(btnId, loading, label = 'Guardar') {
  const btn = document.getElementById(btnId);
  if (!btn) return;
  btn.disabled = loading;
  btn.innerHTML = loading
    ? `<svg class="animate-spin w-4 h-4 mx-auto" viewBox="0 0 24 24" fill="none">
         <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
         <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
       </svg>`
    : label;
}
