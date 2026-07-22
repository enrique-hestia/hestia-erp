/* ══════════════════════════════════════════════════════════════════════
   CHAT INTERNO — canal general + mensajes directos (DM), por sondeo (polling)
   Requiere: SHEET_ID (core.gs), jsonResponse. Wiring POST en finance.gs.
   Hoja 'Chat' (Fecha | Canal | DeEmail | DeNombre | Mensaje) en el Sheets
   principal. Presencia (quién está en línea) en Script Property CHAT_PRES.
   NO es tiempo real: los mensajes llegan en el siguiente sondeo (~7s).
   ══════════════════════════════════════════════════════════════════════ */
var CHAT_TAB = 'Chat';
var CHAT_LECT_TAB = 'Chat_Lecturas';   // estado de LEÍDO por usuario (sincroniza entre navegadores)

function _chatSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CHAT_TAB);
  if (!sh) { sh = ss.insertSheet(CHAT_TAB); sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Canal', 'DeEmail', 'DeNombre', 'Mensaje']]); }
  return sh;
}

/* ── Estado de LEÍDO por usuario (persistente/servidor) ─────────────────────
   El "hasta cuándo leyó" vivía SOLO en localStorage del navegador → al abrir en
   otra máquina/incógnito todo salía como NO leído. Ahora se guarda por usuario
   en la hoja Chat_Lecturas (Email | MapaJSON {canal:ts} | ActualizadoEn) y viaja
   en chatPoll, así el estado de leído sincroniza entre navegadores. */
function _chatLecturasSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CHAT_LECT_TAB);
  if (!sh) { sh = ss.insertSheet(CHAT_LECT_TAB); sh.getRange(1, 1, 1, 3).setValues([['Email', 'Mapa', 'ActualizadoEn']]); }
  return sh;
}
function _chatGetLecturas(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return {};
  try {
    var sh = _chatLecturasSheet(); var lr = sh.getLastRow();
    if (lr < 2) return {};
    var vals = sh.getRange(2, 1, lr - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').trim().toLowerCase() === email) {
        try { return JSON.parse(vals[i][1] || '{}') || {}; } catch(e) { return {}; }
      }
    }
  } catch(e) {}
  return {};
}
/* Marca canal(es) como leídos hasta cierto ts. body: {email, canal, ts} o {email, mapa:{canal:ts}}.
   Nunca RETROCEDE un ts. Si nada avanza, NO toma el lock ni escribe (evita
   amplificar el candado global en cada sondeo con la vista abierta). */
function chatMarkRead(body) {
  try {
    body = body || {};
    var email = String(body.email || '').trim().toLowerCase();
    if (!email) return { ok: false, error: 'email requerido' };
    var updates = (body.mapa && typeof body.mapa === 'object') ? body.mapa : {};
    if (body.canal) updates[String(body.canal)] = Number(body.ts) || 0;   // sin fallback a "ahora": un 0 no marca de más
    // Estado actual (lectura barata, sin lock) para decidir si hay algo que avanzar.
    var actual = _chatGetLecturas(email), cambia = false, c;
    for (c in updates) { if (updates.hasOwnProperty(c) && (Number(updates[c]) || 0) > (Number(actual[c]) || 0)) { cambia = true; break; } }
    if (!cambia) return { ok: true, lecturas: actual, sinCambio: true };   // nada que persistir → sin lock ni escritura
    var lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch(e) { return { ok: false, error: 'lock' }; }
    try {
      var sh = _chatLecturasSheet(); var lr = sh.getLastRow();
      var rowIdx = -1, mapa = {};
      if (lr >= 2) {
        var vals = sh.getRange(2, 1, lr - 1, 2).getValues();
        for (var i = 0; i < vals.length; i++) {
          if (String(vals[i][0] || '').trim().toLowerCase() === email) {
            rowIdx = i + 2; try { mapa = JSON.parse(vals[i][1] || '{}') || {}; } catch(e) { mapa = {}; }
            break;
          }
        }
      }
      var toco = false;
      for (c in updates) {
        if (!updates.hasOwnProperty(c)) continue;
        var t = Number(updates[c]) || 0;
        if (t > (Number(mapa[c]) || 0)) { mapa[c] = t; toco = true; }
      }
      if (toco) {
        var payload = [email, JSON.stringify(mapa), new Date()];
        if (rowIdx > -1) sh.getRange(rowIdx, 1, 1, 3).setValues([payload]);
        else sh.appendRow(payload);
      }
      return { ok: true, lecturas: mapa };
    } finally { try { lock.releaseLock(); } catch(e) {} }
  } catch (ex) { return { ok: false, error: ex.message }; }
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

// Adjuntar archivo: guarda en Drive (carpeta Chat_Adjuntos), lo comparte por link y
// publica un mensaje con el enlace. Límite ~9 MB. Devuelve la URL.
function chatUpload(body) {
  try {
    var email = String((body && body.email) || '').trim().toLowerCase();
    var nombre = String((body && body.nombre) || '').trim() || email;
    var canal = String((body && body.canal) || 'general').trim() || 'general';
    var name = String((body && body.filename) || 'archivo').trim() || 'archivo';
    var mime = String((body && body.mimeType) || 'application/octet-stream');
    var data = String((body && body.dataBase64) || '').replace(/^data:[^;]+;base64,/, '');
    if (!email) return { ok: false, error: 'Sin usuario.' };
    if (!data) return { ok: false, error: 'Archivo vacío.' };
    var bytes = Utilities.base64Decode(data);
    if (bytes.length > 9 * 1024 * 1024) return { ok: false, error: 'Archivo muy grande (máx 9 MB).' };
    var blob = Utilities.newBlob(bytes, mime, name);
    var it = DriveApp.getFoldersByName('Chat_Adjuntos');
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder('Chat_Adjuntos');
    var f = folder.createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    var url = f.getUrl();
    var sent = chatSend({ email: email, nombre: nombre, canal: canal, mensaje: '📎 ' + name + ' | ' + url });
    return { ok: true, url: url, name: name, ts: sent && sent.ts };
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
// Devuelve también el `alias` de cada quien: es lo que el chat pinta en pantalla
// (el nombre completo sigue viajando en `nombre` y sin tocar en la hoja Usuarios).
function _chatDirectorio() {
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID); var sh = ss.getSheetByName('Usuarios'); if (!sh) return out;
    var d = sh.getDataRange().getValues(); var h = d[0].map(function (c) { return String(c).trim().toLowerCase(); });
    var eI = h.indexOf('email'), nI = h.indexOf('nombre'), aI = h.indexOf('activo'), lI = h.indexOf('alias');
    if (eI < 0) return out;
    for (var i = 1; i < d.length; i++) {
      var em = String(d[i][eI] || '').trim(); if (!em) continue;
      var act = aI > -1 ? d[i][aI] : true; if (act === false || String(act).toUpperCase() === 'FALSE') continue;
      var nom = String(d[i][nI > -1 ? nI : 0] || em).trim();
      var ali = lI > -1 ? String(d[i][lI] || '').trim() : '';
      out.push({
        email: em.toLowerCase(), nombre: nom,
        alias: (typeof _aliasFor === 'function') ? _aliasFor(ali, nom, em) : (ali || nom || em)
      });
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
    // Avisos generales a caballo del mismo sondeo (cero timers nuevos en el
    // cliente). Si avisos.gs no está desplegado, el chat sigue como si nada.
    try { if (typeof avisosListActivos === 'function') { var _av = avisosListActivos(); if (_av && _av.ok) res.avisos = _av.avisos; } } catch(e){}
    // Estado de LEÍDO sincronizado entre navegadores: el poll persiste el canal
    // que el usuario tiene abierto y devuelve su mapa completo para sembrar el
    // "hasta cuándo leyó" en cualquier navegador (no depende de localStorage).
    if (body.leerCanal) { try { chatMarkRead({ email: email, canal: body.leerCanal, ts: Number(body.leerTs) || Date.now() }); } catch(e){} }
    res.lecturas = _chatGetLecturas(email);
    return res;
  } catch (ex) { return { ok: false, error: ex.message }; }
}
