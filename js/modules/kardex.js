/**
 * kardex.js
 * Módulo Kardex — Fase 1
 *
 * Este archivo está preparado para recibir la lógica completa del módulo.
 * La función initKardex es llamada por router.js cuando se navega a /kardex.
 */

export async function initKardex(session) {
  const container = document.getElementById('kardex-root');
  if (!container) return;

  // La implementación completa del Kardex se desarrolla en la Fase 1.
  // Por ahora solo muestra el estado del módulo.
  container.innerHTML = `
    <div class="flex flex-col items-center justify-center py-20 text-center gap-4">
      <div class="w-16 h-16 rounded-2xl bg-accent/10 flex items-center justify-center">
        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#2196F3" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/>
          <polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>
        </svg>
      </div>
      <div>
        <h2 class="text-lg font-semibold text-gray-800">Kardex</h2>
        <p class="text-sm text-gray-500 mt-1 max-w-xs">
          Módulo en construcción. Se implementará en la Fase 1.
        </p>
      </div>
      <span class="badge badge-blue">Próximamente</span>
    </div>
  `;
}
