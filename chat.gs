/* ══════════════════════════════════════════════════════════════════════
   CHAT INTERNO — canal general + mensajes directos (DM), por sondeo (polling)
   Requiere: SHEET_ID (core.gs), jsonResponse. Wiring POST en finance.gs.
   Hoja 'Chat' (Fecha | Canal | DeEmail | DeNombre | Mensaje) en el Sheets
   principal. Presencia (quién está en línea) en Script Property CHAT_PRES.
   NO es tiempo real: los mensajes llegan en el siguiente sondeo (~7s).
   ══════════════════════════════════════════════════════════════════════ */
var CHAT_TAB = 'Chat';

function _chatSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CHAT_TAB);
  if (!sh) { sh = ss.insertSheet(CHAT_TAB); sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Canal', 'DeEmail', 'DeNombre', 'Mensaje']]); }
  return sh;
}

// Enviar un mensaje (canal 'general' o 'dm:emailA|emailB' ordenado).
function chatSend(body) {
  try {
    var email = String((body && body.email) || '').trim().toLowerCase();
    var nombre = String((body && body.nombre) || '').trim() || email;
    var canal = String((body && body.canal) || 'general').trim() || 'general';
    var msg = String((body && body.mensaje) || '').trim();
    if (!email) return { ok: false, error: 'Sin usuario.' };
    if (!msg) return { ok: false, error: 'Mensaje vacío.' };
    if (msg.length > 2000) msg = msg.substring(0, 2000);
    var lock = LockService.getScriptLock(); try { lock.waitLock(5000); } catch (e) {}
    var sh = _chatSheet(); var now = new Date();
    sh.appendRow([now, canal, email, nombre, msg]);
    try { lock.releaseLock(); } catch (e) {}
    return { ok: true, ts: now.getTime() };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Registra el "latido" del usuario y devuelve la lista de conectados (ping < 90s).
function _chatPresence(email, nombre) {
  var out = [];
  try {
    var lock = LockService.getScriptLock(); try { lock.waitLock(4000); } catch (e) {}
    var p = PropertiesService.getScriptProperties();
    var raw = p.getProperty('CHAT_PRES'); var m = raw ? JSON.parse(raw) : {};
    var now = Date.now();
    if (email) m[email] = { n: nombre || email, t: now };
    Object.keys(m).forEach(function (k) { if (now - (m[k].t || 0) > 90000) delete m[k]; });
    p.setProperty('CHAT_PRES', JSON.stringify(m));
    try { lock.releaseLock(); } catch (e) {}
    Object.keys(m).forEach(function (k) { out.push({ email: k, nombre: m[k].n || k }); });
  } catch (e) {}
  return out;
}

// Directorio de usuarios activos (para iniciar un DM con quien no esté en línea).
function _chatDirectorio() {
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID); var sh = ss.getSheetByName('Usuarios'); if (!sh) return out;
    var d = sh.getDataRange().getValues(); var h = d[0].map(function (c) { return String(c).trim().toLowerCase(); });
    var eI = h.indexOf('email'), nI = h.indexOf('nombre'), aI = h.indexOf('activo');
    if (eI < 0) return out;
    for (var i = 1; i < d.length; i++) {
      var em = String(d[i][eI] || '').trim(); if (!em) continue;
      var act = aI > -1 ? d[i][aI] : true; if (act === false || String(act).toUpperCase() === 'FALSE') continue;
      out.push({ email: em.toLowerCase(), nombre: String(d[i][nI > -1 ? nI : 0] || em).trim() });
    }
  } catch (e) {}
  return out;
}

// Sondeo: registra presencia + devuelve mensajes visibles nuevos (ts > sinceTs) + conectados.
function chatPoll(body) {
  try {
    body = body || {};
    var email = String(body.email || '').trim().toLowerCase();
    var nombre = String(body.nombre || '').trim() || email;
    var since = Number(body.sinceTs) || 0;
    var online = _chatPresence(email, nombre);
    var sh = _chatSheet(); var lr = sh.getLastRow(); var mensajes = [];
    if (lr > 1) {
      var startRow = Math.max(2, lr - 600); // últimas ~600 filas
      var vals = sh.getRange(startRow, 1, lr - startRow + 1, 5).getValues();
      for (var i = 0; i < vals.length; i++) {
        var f = vals[i][0]; var ts = (f instanceof Date) ? f.getTime() : Date.parse(f); if (!ts) continue;
        if (since && ts <= since) continue;
        var canal = String(vals[i][1] || 'general').trim();
        var vis = canal === 'general';
        if (!vis && canal.indexOf('dm:') === 0) { var parts = canal.substring(3).split('|'); vis = parts.indexOf(email) > -1; }
        if (!vis) continue;
        mensajes.push({ ts: ts, canal: canal, deEmail: String(vals[i][2] || ''), deNombre: String(vals[i][3] || ''), mensaje: String(vals[i][4] || '') });
      }
    }
    var res = { ok: true, mensajes: mensajes, online: online, serverTs: Date.now() };
    if (!since || body.wantDir) res.dir = _chatDirectorio(); // directorio solo en el primer sondeo
    return res;
  } catch (ex) { return { ok: false, error: ex.message }; }
}
