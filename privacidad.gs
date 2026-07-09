/**
 * privacidad.gs — Enmascaramiento de datos sensibles (RBAC)
 *
 * Objetivo (pedido de los socios): quien NO tenga el permiso `ver_datos_sensibles`
 * no debe poder ver NI DEDUCIR nombres de pacientes, empleados o proveedores.
 * Solo el número de operación es suficiente.
 *
 * Se resuelve en el SERVIDOR: los nombres ni siquiera salen del backend para
 * usuarios no autorizados. Los TOTALES y números NO cambian — solo se ocultan
 * las etiquetas/nombres.
 *
 * Flujo:
 *  - core.gs (doGet), tras validar el token, llama _privSet(ss, currentUser).
 *  - Las funciones de lectura (finance/summary/board) consultan _privVer() y,
 *    si es false, aplican las máscaras al construir cada fila.
 *
 * Requiere: core.gs (hoja Roles con columna `permisos_operativos`).
 */

// Request-scoped. Default TRUE = comportamiento normal (ver todo).
// core.gs lo baja a FALSE para usuarios sin el permiso.
var _REQ_VER_SENSIBLES = true;

// ¿El usuario de ESTA petición puede ver datos sensibles?
function _privVer() { return _REQ_VER_SENSIBLES !== false; }

// Fija el flag a partir del rol/permisos del usuario autenticado.
function _privSet(ss, currentUser) {
  try { _REQ_VER_SENSIBLES = _privResuelve(ss, currentUser); }
  catch (e) { _REQ_VER_SENSIBLES = false; }   // fail-secure: ante duda, ocultar
  return _REQ_VER_SENSIBLES;
}

function _privResuelve(ss, currentUser) {
  // Sin módulo de usuarios (modo abierto/dev) → ver todo.
  if (!currentUser) return true;
  var rol = String(currentUser.rol || '').toLowerCase();
  if (rol === 'admin' || rol === 'director') return true;   // siempre ven todo
  var perms = _privPermisosDeRol(ss, currentUser.rol);
  if (perms.length) {   // el rol YA está configurado con permisos operativos → honrarlos
    return perms.indexOf('*') !== -1 || perms.indexOf('ver_datos_sensibles') !== -1;
  }
  // Rol sin permisos operativos configurados → NO romper a los usuarios actuales:
  // solo los roles restringidos se enmascaran por defecto; el resto ve.
  var restringidos = { socio:1, viewer:1, invitado:1, consulta:1, externo:1 };
  return !restringidos[rol];
}

// Lee la columna `permisos_operativos` de la hoja Roles para un rol.
function _privPermisosDeRol(ss, rol) {
  try {
    var sh = ss.getSheetByName('Roles'); if (!sh) return [];
    var data = sh.getDataRange().getValues(); if (data.length < 2) return [];
    var h = data[0].map(function (c) { return String(c).trim().toLowerCase(); });
    var rI = h.indexOf('rol');
    var pI = h.indexOf('permisos_operativos'); if (pI < 0) pI = h.indexOf('permisos');
    if (rI < 0 || pI < 0) return [];
    var target = String(rol || '').trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][rI]).trim().toLowerCase() === target) {
        return String(data[i][pI] || '').split(',').map(function (v) { return v.trim(); }).filter(Boolean);
      }
    }
    return [];
  } catch (e) { return []; }
}

// ── Máscaras concretas ──────────────────────────────────────────────
// Paciente → solo el número de operación (nada de nombre ni iniciales).
function _privPaciente(op) {
  var o = String(op == null ? '' : op).trim();
  return o ? ('OP-' + o) : 'OP —';
}
// Proveedor / empleado / nombre libre → se oculta por completo.
// Al devolver '' , el agrupador de reportes cae al subtipo/categoría (colapsa
// nómina y proveedores a su total, sin exponer nombres).
function _privNombre() { return ''; }
// Texto libre (concepto, observaciones) que puede contener nombres → genérico.
function _privTextoLibre(fallback) { return String(fallback || ''); }
// Referencia/observación bancaria (contiene "Px Nombre", "LABORATORIO · VIANTO"…)
// → se oculta si el usuario no está autorizado.
function _privRef(texto) { return _privVer() ? String(texto == null ? '' : texto) : ''; }

/* Crea (una sola vez) el rol "Socio": solo lectura, con las vistas operativas y de
   PII bloqueadas. El enmascarado de datos ya lo garantiza el backend aunque las
   vistas no se bloqueen; esto es la capa de UX. Ejecutar desde el editor de Apps
   Script. Ajusta las vistas bloqueadas en Panel de Control → Gestión de Roles. */
function setupRolSocio() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Roles');
  if (!sh) { sh = ss.insertSheet('Roles'); sh.appendRow(['rol', 'vistas_bloqueadas', 'solo_lectura', 'permisos_operativos', 'descripcion']); }
  var data = sh.getDataRange().getValues();
  var h = data[0].map(function (c) { return String(c).trim().toLowerCase(); });
  var rI = h.indexOf('rol');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][rI]).trim().toLowerCase() === 'socio') return { ok: true, yaExiste: true };
  }
  // Vistas típicas a bloquear (ajústalas en Gestión de Roles si tus IDs difieren).
  var bloq = 'pacientes,fin-ingresos,fin-egresos,fin-cxp,conciliacion,proveedores,caja-chica,comprobantes,gestion-roles,panel-control,chat';
  var row = h.map(function (col) {
    if (col === 'rol') return 'Socio';
    if (col.indexOf('vistas') > -1) return bloq;
    if (col.indexOf('solo') > -1) return true;
    if (col.indexOf('permisos') > -1) return 'export_data';
    if (col.indexOf('desc') > -1) return 'Socios: reportes financieros, sin datos sensibles';
    return '';
  });
  sh.appendRow(row);
  return { ok: true, creado: true, vistasBloqueadas: bloq };
}
