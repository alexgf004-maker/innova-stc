/**
 * home.js — Pantalla de inicio
 * Admin: dashboard ejecutivo
 * Coordinadora/Campo: accesos rápidos
 */

import {
  collection, getDocs, query, where, orderBy
} from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';

// Simple in-memory cache — clears on page reload
const _cache = {};
function cached(key, ttlMs, fetcher) {
  const now = Date.now();
  if (_cache[key] && (now - _cache[key].ts) < ttlMs) return Promise.resolve(_cache[key].data);
  return fetcher().then(function(data) { _cache[key] = { data, ts: now }; return data; });
}

export async function initHome(session) {
  // Greeting for all roles
  const hour  = new Date().getHours();
  const greet = hour < 12 ? 'Buenos días' : hour < 19 ? 'Buenas tardes' : 'Buenas noches';
  const greetEl = document.getElementById('home-greeting');
  if (greetEl) greetEl.textContent = greet + ', ' + session.displayName.split(' ')[0];

  if (session.role === 'admin') {
    await initDashboardAdmin(session);
  } else {
    initHomeSimple(session);
  }
}

function initHomeSimple(session) {
  const el = document.getElementById('home-modules');
  if (!el) return;

  const isCampo = session.role === 'campo';
  const area    = session.asignacionActual?.area || (session.usuarioOperativoAsignado ? 'OTC' : null);
  const destino = session.asignacionActual?.destino || session.usuarioOperativoAsignado || null;
  const color   = 'linear-gradient(135deg,#0f1f3d,#1a3a6b)';

  el.innerHTML =
    '<div class="space-y-4">' +
      // Asignación actual
      (isCampo && area ?
        '<div style="background:' + color + ';border-radius:14px;padding:16px;color:white">' +
          '<p style="font-size:11px;font-weight:600;opacity:.8;text-transform:uppercase;letter-spacing:.05em">Trabajando como</p>' +
          '<p style="font-size:22px;font-weight:900;margin-top:2px">' + (destino || '—') + '</p>' +
          '<p style="font-size:12px;opacity:.7;margin-top:2px">Área: ' + area + '</p>' +
        '</div>'
      : '') +
      // Accesos rápidos
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">' +
        (session.role === 'coordinadora' ?
          ['<a data-route="/kardex" style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:14px;cursor:pointer;text-decoration:none;display:block"><p style="font-size:13px;font-weight:600;color:#111827">Kardex</p><p style="font-size:11px;color:#9ca3af;margin-top:2px">Bodega y materiales</p></a>',
           '<a data-route="/cm" style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:14px;cursor:pointer;text-decoration:none;display:block"><p style="font-size:13px;font-weight:600;color:#111827">Cambios</p><p style="font-size:11px;color:#9ca3af;margin-top:2px">Gestión de órdenes</p></a>'].join('')
        : area === 'CAMBIOS' ?
          '<a data-route="/cm" style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:14px;cursor:pointer;text-decoration:none;display:block"><p style="font-size:13px;font-weight:600;color:#111827">Mis órdenes</p><p style="font-size:11px;color:#9ca3af;margin-top:2px">Cambio de medidores</p></a>'
        : area === 'OTC' ?
          '<a data-route="/otc" style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:14px;cursor:pointer;text-decoration:none;display:block"><p style="font-size:13px;font-weight:600;color:#111827">OTC</p><p style="font-size:11px;color:#9ca3af;margin-top:2px">Mis órdenes y bodega</p></a>'
        : '<p style="font-size:13px;color:#9ca3af;grid-column:1/-1;text-align:center;padding:20px">Sin área asignada. Contacta a administración.</p>') +
      '</div>' +
    '</div>';
}

async function initDashboardAdmin(session) {
  const db = window.__firebase.db;

  const dashboard = document.getElementById('home-modules');
  if (!dashboard) return;
  dashboard.innerHTML = '<div class="text-center py-8 text-gray-400 text-sm">Cargando dashboard...</div>';

  try {
    // Load all data in parallel
    const [snapUsers, snapOrdenes, snapItems, snapSolicitudes] = await Promise.all([
      getDocs(collection(db, 'users')),
      getDocs(collection(db, 'cambios_ordenes')),
      getDocs(collection(db, 'kardex/inventario/items')),
      getDocs(query(collection(db, 'solicitudes_material'), where('estado', '==', 'pendiente'))),
    ]);

    const users    = snapUsers.docs.map(d => Object.assign({ id: d.id }, d.data())).filter(u => u.active);
    const ordenes  = snapOrdenes.docs.map(d => Object.assign({ id: d.id }, d.data()));
    const items    = snapItems.docs.map(d => Object.assign({ id: d.id }, d.data()));

    // Today boundaries
    const hoy    = new Date(); hoy.setHours(0,0,0,0);
    const manana = new Date(hoy); manana.setDate(manana.getDate()+1);
    function esHoy(ts) {
      if (!ts) return false;
      const d = ts.toDate ? ts.toDate() : new Date(ts);
      return d >= hoy && d < manana;
    }

    // Corte del 20
    const ahora  = new Date();
    const corte  = new Date(ahora.getFullYear(), ahora.getMonth(), 20);
    if (ahora > corte) corte.setMonth(corte.getMonth() + 1);
    const diasCorte = Math.ceil((corte - ahora) / (1000*60*60*24));

    // Campo activo hoy — users con asignacionActual
    const campoCambios = users.filter(u => u.role === 'campo' && u.asignacionActual?.area === 'CAMBIOS');
    const campoOTC     = users.filter(u => u.role === 'campo' && u.asignacionActual?.area === 'OTC');

    // Agrupar cambios por pareja
    const PAREJAS = ['Pareja 1','Pareja 2','Pareja 3','Pareja 4'];
    const PAREJA_COLORS = { 'Pareja 1':'#1B4F8A','Pareja 2':'#EA580C','Pareja 3':'#7C3AED','Pareja 4':'#DB2777' };
    const parejaStats = {};
    PAREJAS.forEach(p => {
      const miembros = campoCambios.filter(u => u.asignacionActual?.destino === p);
      const ords     = ordenes.filter(o => o.pareja === p && o.estadoCampo !== 'aprobada');
      const hechasHoy = ords.filter(o => o.estadoCampo === 'hecha' && esHoy(o.fechaHecha));
      const visitasHoy = ords.filter(o => o.estadoCampo === 'visita' && esHoy(o.fechaVisita));
      parejaStats[p] = { miembros, total: ords.length, hechasHoy: hechasHoy.length, visitasHoy: visitasHoy.length };
    });

    // Cambios global stats
    const totalOrdenes    = ordenes.filter(o => o.estadoCampo !== 'aprobada').length;
    const realizadasHoy   = ordenes.filter(o => o.estadoCampo === 'hecha' && esHoy(o.fechaHecha)).length;
    const sinActualizar   = ordenes.filter(o => o.estadoCampo === 'hecha' && !o.actualizadaDelsur).length;
    const pendConfirm     = ordenes.filter(o => o.estadoCampo === 'hecha').length;

    // Bodega alerts
    const itemsBajos = items.filter(i => i.stockActual !== undefined && i.stockActual <= (i.stockMinimo || 5));
    const solicsPend = snapSolicitudes.size;

    // Corte urgency
    const corteColor  = diasCorte <= 3 ? '#C62828' : diasCorte <= 7 ? '#B45309' : '#166534';
    const corteBg     = diasCorte <= 3 ? '#FEF2F2' : diasCorte <= 7 ? '#FEF3C7' : '#F0FDF4';
    const corteBorder = diasCorte <= 3 ? '#FECACA' : diasCorte <= 7 ? '#FDE68A' : '#BBF7D0';
    const fmtCorte    = corte.toLocaleDateString('es-SV', { day:'2-digit', month:'long' });

    function statCard(val, label, color, bg) {
      return '<div style="background:' + bg + ';border-radius:12px;padding:12px;text-align:center">' +
        '<p style="font-size:22px;font-weight:900;color:' + color + ';line-height:1">' + val + '</p>' +
        '<p style="font-size:11px;color:#6b7280;margin-top:3px">' + label + '</p>' +
      '</div>';
    }

    function miembroChip(u) {
      return '<span style="font-size:11px;background:#f3f4f6;border-radius:20px;padding:2px 8px;color:#374151">' +
        (u.displayName || u.username) +
      '</span>';
    }

    dashboard.innerHTML =
      '<div class="space-y-4">' +

        // ── Corte del 20 ──
        '<div style="background:' + corteBg + ';border:1.5px solid ' + corteBorder + ';border-radius:12px;padding:14px;display:flex;align-items:center;justify-content:space-between">' +
          '<div>' +
            '<p style="font-size:13px;font-weight:700;color:' + corteColor + '">Corte: ' + fmtCorte + '</p>' +
            '<p style="font-size:11px;color:' + corteColor + ';opacity:.8">' + (diasCorte === 0 ? '¡Hoy es el corte!' : diasCorte === 1 ? 'Mañana es el corte' : 'Faltan ' + diasCorte + ' días') + '</p>' +
          '</div>' +
          '<div style="text-align:right">' +
            '<p style="font-size:26px;font-weight:900;color:' + corteColor + ';line-height:1">' + sinActualizar + '</p>' +
            '<p style="font-size:10px;color:' + corteColor + ';opacity:.8">sin actualizar</p>' +
          '</div>' +
        '</div>' +

        // ── Stats cambios ──
        '<div>' +
          '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">' +
            '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em">Cambios de medidores · Hoy</p>' +
            '<a data-route="/cm" style="font-size:11px;color:#0F766E;font-weight:600;text-decoration:none">Ver más →</a>' +
          '</div>' +
          '<a data-route="/cm" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;text-decoration:none">' +
            statCard(realizadasHoy, 'Realizadas hoy', '#166534', '#F0FDF4') +
            statCard(pendConfirm, 'Pendientes aprobación', pendConfirm > 0 ? '#B45309' : '#166534', pendConfirm > 0 ? '#FEF3C7' : '#F0FDF4') +
            statCard(totalOrdenes, 'Total activas', '#1B4F8A', '#EFF6FF') +
          '</a>' +
        '</div>' +

        // ── Parejas activas ──
        (campoCambios.length > 0 ?
          '<div>' +
            '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">' +
            '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em">Parejas activas · Cambios (' + campoCambios.length + ' personas)</p>' +
            '<a data-route="/cm" style="font-size:11px;color:#0F766E;font-weight:600;text-decoration:none">Seguimiento →</a>' +
          '</div>' +
            '<div class="space-y-2">' +
              PAREJAS.filter(p => parejaStats[p].miembros.length > 0).map(p => {
                const s = parejaStats[p];
                const color = PAREJA_COLORS[p];
                const pct = Math.min(100, Math.round(((s.hechasHoy + s.visitasHoy) / 15) * 100));
                return '<a data-route="/cm" style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:12px;display:block;cursor:pointer;text-decoration:none">' +
                  '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">' +
                    '<div style="display:flex;align-items:center;gap:6px">' +
                      '<span style="width:10px;height:10px;border-radius:50%;background:' + color + ';display:inline-block"></span>' +
                      '<p style="font-size:13px;font-weight:700;color:#111827">' + p + '</p>' +
                    '</div>' +
                    '<div style="display:flex;align-items:center;gap:8px">' +
                      (s.sinActualizar > 0 ? '<span style="font-size:10px;font-weight:700;padding:2px 7px;border-radius:20px;background:#FEF3C7;color:#B45309">⚠ ' + s.sinActualizar + ' sin actualizar</span>' : '') +
                      '<p style="font-size:11px;color:#6b7280">' + (s.hechasHoy + s.visitasHoy) + '/15</p>' +
                    '</div>' +
                  '</div>' +
                  '<div style="display:flex;gap:4px;flex-wrap:wrap;margin-bottom:8px">' +
                    s.miembros.map(miembroChip).join('') +
                  '</div>' +
                  '<div style="height:6px;background:#f3f4f6;border-radius:3px;overflow:hidden;margin-bottom:6px">' +
                    '<div style="height:100%;width:' + pct + '%;background:' + color + ';border-radius:3px;transition:width .3s"></div>' +
                  '</div>' +
                  '<div style="display:flex;gap:12px">' +
                    '<p style="font-size:11px;color:#166534">✓ ' + s.hechasHoy + ' realizadas</p>' +
                    (s.visitasHoy ? '<p style="font-size:11px;color:#374151">👁 ' + s.visitasHoy + ' visitas</p>' : '') +
                    '<p style="font-size:11px;color:#9ca3af">' + s.total + ' asignadas</p>' +
                  '</div>' +
                '</a>';
              }).join('') +
            '</div>' +
          '</div>'
        : '<div style="background:#f9fafb;border-radius:12px;padding:14px;text-align:center"><p style="font-size:13px;color:#9ca3af">Sin personal asignado a Cambios hoy</p></div>') +

        // ── OTC ──
        '<div>' +
          '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">' +
            '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em">OTC' + (campoOTC.length ? ' (' + campoOTC.length + ' personas)' : '') + '</p>' +
            '<a data-route="/kardex" style="font-size:11px;color:#1B4F8A;font-weight:600;text-decoration:none">Ver Kardex →</a>' +
          '</div>' +
          (campoOTC.length > 0 ?
            '<div style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:12px">' +
              '<div style="display:flex;gap:4px;flex-wrap:wrap;margin-bottom:6px">' +
                campoOTC.map(u => {
                  const dest = u.asignacionActual?.destino || '—';
                  return '<div style="background:#EFF6FF;border-radius:8px;padding:4px 10px">' +
                    '<p style="font-size:11px;font-weight:600;color:#1B4F8A">' + (u.displayName || u.username) + '</p>' +
                    '<p style="font-size:10px;color:#6b7280">' + dest + '</p>' +
                  '</div>';
                }).join('') +
              '</div>' +
              '<p style="font-size:11px;color:#9ca3af">Control de órdenes OTC próximamente</p>' +
            '</div>'
          : '<div style="background:#f9fafb;border-radius:12px;padding:14px;text-align:center"><p style="font-size:13px;color:#9ca3af">Sin personal asignado a OTC hoy</p></div>') +
        '</div>' +

        // ── Bodega ──
        '<div>' +
          '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">' +
            '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em">Bodega</p>' +
            '<a data-route="/kardex" style="font-size:11px;color:#1B4F8A;font-weight:600;text-decoration:none">Ver Kardex →</a>' +
          '</div>' +
          '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px">' +
            statCard(items.length, 'Items en inventario', '#1B4F8A', '#EFF6FF') +
            statCard(solicsPend, 'Solicitudes pendientes', solicsPend > 0 ? '#B45309' : '#166534', solicsPend > 0 ? '#FEF3C7' : '#F0FDF4') +
            statCard(itemsBajos.length, 'Stock bajo', itemsBajos.length > 0 ? '#C62828' : '#166534', itemsBajos.length > 0 ? '#FEF2F2' : '#F0FDF4') +
          '</div>' +
        '</div>' +

        // ── Accesos rápidos ──
        '<div>' +
          '<p style="font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">Accesos rápidos</p>' +
          '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">' +
            '<a data-route="/cm" style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:12px;cursor:pointer;text-decoration:none;display:block">' +
              '<p style="font-size:13px;font-weight:600;color:#111827">Seguimiento CM</p>' +
              '<p style="font-size:11px;color:#9ca3af;margin-top:2px">Ver avance del día</p>' +
            '</a>' +
            '<a data-route="/kardex" style="background:white;border:1px solid #e5e7eb;border-radius:12px;padding:12px;cursor:pointer;text-decoration:none;display:block">' +
              '<p style="font-size:13px;font-weight:600;color:#111827">Kardex</p>' +
              '<p style="font-size:11px;color:#9ca3af;margin-top:2px">Bodega y materiales</p>' +
            '</a>' +
          '</div>' +
        '</div>' +

      '</div>';

  } catch(e) {
    dashboard.innerHTML = '<div class="text-center py-8 text-sm text-red-400">Error al cargar el dashboard.</div>';
    console.error(e);
  }
}
