/* ═══════════════════════════════════════════════════════════════════════════
 *  novedades.gs — CONFIG DEL AVISO DE NOVEDADES  (hoja 'Config_Novedades')
 *
 *  QUÉ RESUELVE
 *  El catálogo de novedades (array NOVEDADES) vive en el HTML, pero QUÉ se anuncia
 *  y A QUIÉN es una decisión de operación, no de código: el director prende el aviso
 *  solo cuando hay algo importante y elige si una novedad es para todos o solo para
 *  quien tenga cierto permiso.
 *
 *  POR QUÉ EN UNA HOJA Y NO EN localStorage
 *  Es una decisión del ADMIN que debe aplicar a TODOS los usuarios. localStorage es
 *  por navegador: el admin prendería el aviso solo para sí mismo y nadie más lo vería
 *  — justo lo contrario de lo que se pide. La hoja es compartida y persistente.
 *  (El "ya lo vi, no me lo muestres otra vez" SÍ es por usuario y se queda en
 *   localStorage del navegador: eso es preferencia personal, no configuración.)
 *
 *  MODELO
 *    Clave | Destacada | Audiencia | Notas
 *    · Clave '__AVISO__'  → INTERRUPTOR MAESTRO. Destacada TRUE/FALSE = el aviso
 *      emergente salta o no. Apagado NO esconde la vista de Novedades: solo evita
 *      que salte solo (se sigue consultando desde el menú a propósito).
 *    · Cualquier otra Clave → una entrada del array NOVEDADES ('build|titulo', o su
 *      id si lo trae). Destacada = se anuncia. Audiencia = '' (todos) o la clave de
 *      un permiso de OP_PERMS_CATALOG (solo quien lo tenga la ve).
 *
 *  La hoja SOBRESCRIBE los valores por defecto del array. Si no hay hoja / no está
 *  redesplegado, el frontend usa los defaults del código y sigue funcionando.
 * ═══════════════════════════════════════════════════════════════════════════ */

var NOV_CFG_TAB = 'Config_Novedades';
var NOV_CFG_MASTER = '__AVISO__';

function _novBool(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE' || String(v).trim() === '1';
}

/* Solo admin/director pueden tocar la config del aviso: decide qué ve TODA la
   empresa. Se valida en el SERVIDOR — esconder el tab en el frontend es cosmético
   (cualquiera puede mandar el POST a mano). */
function _novEsAdmin(token) {
  try {
    if (typeof verifyToken !== 'function') return false;
    var email = verifyToken(token || '');
    if (!email) return false;
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var u = (typeof getUserRow === 'function') ? getUserRow(ss, email) : null;
    if (!u) return false;
    var rol = String(u.rol || '').toLowerCase();
    return rol === 'admin' || rol === 'director';
  } catch (ex) { return false; }
}

function setupNovedadesCfg() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(NOV_CFG_TAB);
  if (!sh) sh = ss.insertSheet(NOV_CFG_TAB);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['Clave', 'Destacada', 'Audiencia', 'Notas']]);
  sh.getRange(2, 1, 1, 4).setValues([[NOV_CFG_MASTER, false, '',
    'Interruptor maestro del aviso emergente. FALSE = la vista Novedades sigue disponible en el menu, pero no salta sola.']]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  return { ok: true };
}

/* Lectura ABIERTA a cualquier usuario con sesión: el filtro de audiencia se aplica
   en el frontend con hasPermission, y para eso todos necesitan leer la config.
   No expone nada sensible: son banderas de presentación. */
function readNovedadesCfg() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(NOV_CFG_TAB);
    // Sin hoja aún: no es error. El frontend cae a los defaults del código
    // (nada destacado → no salta nada), que es el estado seguro.
    if (!sh) return { ok: true, _setup: false, avisoActivo: false, entradas: {} };
    var raw = sh.getDataRange().getValues();
    var out = { ok: true, _setup: true, avisoActivo: false, entradas: {} };
    for (var i = 1; i < raw.length; i++) {
      var k = String(raw[i][0] || '').trim();
      if (!k) continue;
      if (k === NOV_CFG_MASTER) { out.avisoActivo = _novBool(raw[i][1]); continue; }
      out.entradas[k] = { destacada: _novBool(raw[i][1]), audiencia: String(raw[i][2] || '').trim() };
    }
    return out;
  } catch (ex) {
    return { ok: false, error: ex.message, avisoActivo: false, entradas: {} };
  }
}

/* body: { token, avisoActivo:bool, entradas:[{clave, destacada, audiencia}] }
   Reescribe la hoja completa (la config es chica y siempre se manda entera desde el
   panel; así no quedan filas huérfanas de novedades que ya se borraron del array). */
function saveNovedadesCfg(body) {
  try {
    body = body || {};
    if (!_novEsAdmin(body.token || '')) {
      return { ok: false, error: 'Solo un administrador o director puede configurar el aviso de novedades.' };
    }
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(NOV_CFG_TAB);
    if (!sh) { setupNovedadesCfg(); sh = ss.getSheetByName(NOV_CFG_TAB); }
    sh.clear();
    sh.getRange(1, 1, 1, 4).setValues([['Clave', 'Destacada', 'Audiencia', 'Notas']]);
    var rows = [[NOV_CFG_MASTER, !!body.avisoActivo, '',
      'Interruptor maestro del aviso emergente. FALSE = la vista Novedades sigue disponible en el menu, pero no salta sola.']];
    (body.entradas || []).forEach(function (e) {
      var k = String((e && e.clave) || '').trim();
      if (!k || k === NOV_CFG_MASTER) return;
      rows.push([k, !!e.destacada, String((e && e.audiencia) || '').trim(), String((e && e.titulo) || '')]);
    });
    sh.getRange(1, 1, rows.length, 4).setValues(rows);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    try {
      if (typeof logAudit === 'function') {
        logAudit(body.usuario || 'sistema', 'Novedades', 'Guardar config aviso', '', '', '',
          'aviso=' + (body.avisoActivo ? 'ON' : 'OFF') + ' · ' + (rows.length - 1) + ' entradas');
      }
    } catch (e) {}
    return { ok: true, guardados: rows.length - 1, avisoActivo: !!body.avisoActivo };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}
