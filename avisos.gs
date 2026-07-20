/* ══════════════════════════════════════════════════════════════
   avisos.gs — Avisos generales del equipo (broadcast)
   --------------------------------------------------------------
   Mensajes para TODO el equipo ("cumpleaños de X", "junta a las 10",
   "meta cumplida"), separados de los pendientes personales (Recordatorios):
   hoja distinta, endpoints distintos, permisos distintos.

   - LEER es abierto a cualquier sesión válida (es un broadcast).
   - PUBLICAR/EDITAR exige el permiso `publicar_avisos` (gate en finance.gs).
   - Baja = soft-delete (Activo=FALSE): deja registro, nunca borra filas.
   - Actualizar toca SOLO columnas presentes en el payload (regla anti-blanqueo).
   - `Anim` guarda la receta JSON de la animación (la genera el frontend, E4).
   Hoja "Avisos" en SHEET_ID. El viaje al cliente va a caballo del sondeo del
   chat (chatPoll adjunta `avisos`) — cero timers nuevos.
   ══════════════════════════════════════════════════════════════ */

var AVISOS_TAB = 'Avisos';
var AVISOS_HEADERS = ['Id','CreadoEn','AutorEmail','AutorNombre','Titulo','Mensaje','Nivel','FechaEvento','Activo','Expira','Anim','Segmento'];

function setupAvisos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(AVISOS_TAB);
  if (!sh) sh = ss.insertSheet(AVISOS_TAB);
  sh.clear();
  sh.getRange(1,1,1,AVISOS_HEADERS.length).setValues([AVISOS_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  sh.setFrozenRows(1);
  [110,160,210,150,240,340,80,110,70,110,260,120].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
  return {ok:true, tab:AVISOS_TAB};
}
function _avisosSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(AVISOS_TAB);
  if (!sh) { setupAvisos(); sh = ss.getSheetByName(AVISOS_TAB); }
  return sh;
}
function _avisosHoy() {
  var d = new Date();
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
function _avisosFmtFecha(v) {
  if (!v) return '';
  if (v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
  return String(v).substring(0,10);
}

/* ── Cumpleaños AUTOMÁTICOS desde Nómina ─────────────────────────────────
   Empleados ACTIVOS con FechaNacimiento (hoja Empleados del libro de Nómina,
   vía readEmpleados de nomina.gs — mismo scope global). Son avisos SINTÉTICOS
   del día: no se escriben en la hoja, se calculan al vuelo con id estable por
   persona+año (así el descarte por-usuario del banner funciona igual).
   Privacidad: solo el DÍA — jamás el año de nacimiento ni la edad.
   Cache 6h por fecha: el sondeo del chat pega cada 7s y no vamos a leer el
   libro de Nómina cada vez. */
function _avisosCumplesHoy(hoy) {
  // v2: cambió el texto del título — la llave nueva invalida el caché viejo (6h)
  var cacheKey = 'avisos_cumples_v2_' + hoy;
  try { var hit = CacheService.getScriptCache().get(cacheKey); if (hit !== null) return JSON.parse(hit); } catch(e){}
  var out = [];
  try {
    if (typeof readEmpleados === 'function') {
      var r = readEmpleados();
      var mmdd = hoy.substring(5);   // 'MM-DD'
      ((r && r.empleados) || []).forEach(function(emp){
        if (!emp.Activo) return;
        var fn = String(emp.FechaNacimiento || '');
        if (fn.length < 10 || fn.substring(5) !== mmdd) return;
        var nombre = String(emp.Nombre || '').trim();
        if (!nombre) return;
        var primer = nombre.split(/\s+/)[0];
        primer = primer.charAt(0).toUpperCase() + primer.slice(1).toLowerCase();
        out.push({
          id: 'CUMPLE-' + String(emp.NumEmpleado || primer).replace(/\s+/g,'') + '-' + hoy.substring(0,4),
          ts: Date.parse(hoy + 'T09:00:00') || Date.now(),
          autorEmail: '', autorNombre: 'VestaOS',
          titulo: 'Feliz Cumpleaños ' + primer + ' 🎂',
          mensaje: 'Felicita a ' + nombre + ' de parte de todo el equipo.',
          nivel: 'exito', fechaEvento: hoy, expira: hoy,
          anim: { tpl: 'cumple', nombre: primer, seed: 7 },
          segmento: 'todos', auto: true
        });
      });
    }
  } catch(e){ /* sin nómina configurada → sin cumpleaños automáticos, sin romper avisos */ }
  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(out), 21600); } catch(e){}
  return out;
}

// Lista los avisos VIGENTES (activos y no expirados), más recientes primero.
// Incluye los cumpleaños automáticos del día (sintéticos, van al frente).
function avisosListActivos() {
  try {
    var sh = _avisosSheet();
    var v = sh.getDataRange().getValues();
    var hoy = _avisosHoy(), out = [];
    for (var i = 1; i < v.length; i++) {
      var r = v[i];
      if (!String(r[0]||'').trim()) continue;
      var activo = (r[8] === true || String(r[8]).toLowerCase() === 'true');
      if (!activo) continue;
      var expira = _avisosFmtFecha(r[9]);
      if (expira && expira < hoy) continue;
      var ts = (r[1] instanceof Date) ? r[1].getTime() : Date.parse(r[1]) || 0;
      var anim = null;
      try { anim = r[10] ? JSON.parse(r[10]) : null; } catch(e){ anim = null; }
      out.push({
        id: String(r[0]), ts: ts,
        autorEmail: String(r[2]||''), autorNombre: String(r[3]||''),
        titulo: String(r[4]||''), mensaje: String(r[5]||''),
        nivel: String(r[6]||'info').toLowerCase(),
        fechaEvento: _avisosFmtFecha(r[7]), expira: expira,
        anim: anim, segmento: String(r[11]||'todos')
      });
    }
    out.sort(function(a,b){ return b.ts - a.ts; });
    var cumples = _avisosCumplesHoy(hoy);
    return { ok:true, hoy:hoy, avisos: cumples.concat(out).slice(0, 50) };
  } catch(ex){ return { ok:false, error:ex.message, avisos:[] }; }
}

// Publica un aviso. `autorEmail` viene del TOKEN (finance.gs), no del cliente.
// El gate de permiso (`publicar_avisos`) vive en el router — defensa en capas.
function avisosPublicar(b, autorEmail) {
  try {
    b = b || {};
    var titulo = String(b.titulo||'').trim();
    if (!titulo) return { ok:false, error:'El aviso necesita un título.' };
    var sh = _avisosSheet();
    var lr = sh.getLastRow(), max = 0;
    if (lr > 1) {
      var ids = sh.getRange(2,1,lr-1,1).getValues();
      for (var i=0;i<ids.length;i++){ var m=String(ids[i][0]||'').match(/(\d+)/); if(m){ var n=parseInt(m[1],10); if(n>max) max=n; } }
    }
    var id = 'AVISO-'+String(max+1).padStart(5,'0');
    var nivel = String(b.nivel||'info').toLowerCase();
    if (nivel!=='info' && nivel!=='exito' && nivel!=='alerta') nivel = 'info';
    var anim = '';
    try { if (b.anim && typeof b.anim === 'object') anim = JSON.stringify(b.anim); } catch(e){}
    sh.appendRow([ id, new Date(), String(autorEmail||'').toLowerCase(), String(b.autorNombre||autorEmail||''),
                   titulo, String(b.mensaje||''), nivel, _avisosFmtFecha(b.fechaEvento), true,
                   _avisosFmtFecha(b.expira), anim, String(b.segmento||'todos') ]);
    return { ok:true, id:id };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

// Actualiza un aviso por Id, IN-PLACE y SOLO las columnas presentes en el
// payload (nunca blanquear lo que no vino). Baja = {activo:false} (soft-delete).
function avisosActualizar(b) {
  try {
    b = b || {};
    var id = String(b.id||'').trim();
    if (!id) return { ok:false, error:'Falta el id del aviso.' };
    var sh = _avisosSheet();
    var v = sh.getDataRange().getValues();
    var rowNum = -1;
    for (var i = 1; i < v.length; i++){ if (String(v[i][0]||'').trim() === id) { rowNum = i+1; break; } }
    if (rowNum < 0) return { ok:false, error:'No se encontró el aviso '+id+'.' };
    if (b.titulo      !== undefined) sh.getRange(rowNum,5).setValue(String(b.titulo||''));
    if (b.mensaje     !== undefined) sh.getRange(rowNum,6).setValue(String(b.mensaje||''));
    if (b.nivel       !== undefined) sh.getRange(rowNum,7).setValue(String(b.nivel||'info').toLowerCase());
    if (b.fechaEvento !== undefined) sh.getRange(rowNum,8).setValue(_avisosFmtFecha(b.fechaEvento));
    if (b.activo      !== undefined) sh.getRange(rowNum,9).setValue(!!b.activo);
    if (b.expira      !== undefined) sh.getRange(rowNum,10).setValue(_avisosFmtFecha(b.expira));
    if (b.anim        !== undefined) {
      var a2 = '';
      try { if (b.anim && typeof b.anim === 'object') a2 = JSON.stringify(b.anim); } catch(e){}
      sh.getRange(rowNum,11).setValue(a2);
    }
    if (b.segmento    !== undefined) sh.getRange(rowNum,12).setValue(String(b.segmento||'todos'));
    return { ok:true, id:id };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
