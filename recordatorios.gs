/* ══════════════════════════════════════════════════════════════
   recordatorios.gs — Agenda personal de recordatorios (por usuario)
   --------------------------------------------------------------
   Cualquier área del sistema puede crear un recordatorio llamando a
   saveRecordatorio (POST action) o, desde el frontend, al helper
   ofrecerRecordatorio({titulo, fecha, origen, referencia, detalle}).
   El día objetivo (y los vencidos) salen como banner al iniciar sesión
   y en la vista "Recordatorios". Una tarea programada manda el correo
   diario con los pendientes de cada usuario.
   Hoja "Recordatorios" en SHEET_ID.
   ══════════════════════════════════════════════════════════════ */

var REC_TAB = 'Recordatorios';
var REC_HEADERS = ['ID','Usuario','Titulo','Detalle','FechaObjetivo','Origen','Referencia','Estado','CreadoEn','CompletadoEn'];

function setupRecordatorios() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(REC_TAB);
  if (!sh) sh = ss.insertSheet(REC_TAB);
  sh.clear();
  sh.getRange(1,1,1,REC_HEADERS.length).setValues([REC_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, REC_HEADERS.length, [110,210,260,300,120,110,150,100,160,160]);
  return {ok:true, tab:REC_TAB};
}
function _recSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(REC_TAB);
  if (!sh) { setupRecordatorios(); sh = ss.getSheetByName(REC_TAB); }
  return sh;
}
function _recToday() {
  var d = new Date();
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
function _recFmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
  return String(v).substring(0,10);
}
function _recRows() {
  var sh = _recSheet();
  var v = sh.getDataRange().getValues();
  var rows = [];
  for (var i=1;i<v.length;i++){
    var r = v[i]; if (!String(r[0]||'').trim()) continue;
    rows.push({
      rowNum:i+1, id:String(r[0]), usuario:String(r[1]||'').toLowerCase().trim(),
      titulo:String(r[2]||''), detalle:String(r[3]||''), fecha:_recFmtDate(r[4]),
      origen:String(r[5]||'manual'), referencia:String(r[6]||''),
      estado:String(r[7]||'pendiente').toLowerCase(), creadoEn:_recFmtDate(r[8])
    });
  }
  return rows;
}

// Lee la agenda del usuario, agrupada (pendientes = hoy + vencidos).
function readRecordatorios(usuario) {
  try {
    usuario = String(usuario||'').toLowerCase().trim();
    var hoy = _recToday();
    var mine = _recRows().filter(function(r){ return !usuario || r.usuario===usuario; });
    var pendientes=[], proximos=[], hechos=[];
    mine.forEach(function(r){
      if (r.estado==='hecho' || r.estado==='descartado') hechos.push(r);
      else if (r.fecha && r.fecha<=hoy) pendientes.push(r);   // vencidos + hoy
      else proximos.push(r);
    });
    pendientes.sort(function(a,b){ return (a.fecha||'').localeCompare(b.fecha||''); });
    proximos.sort(function(a,b){ return (a.fecha||'').localeCompare(b.fecha||''); });
    hechos.sort(function(a,b){ return (b.fecha||'').localeCompare(a.fecha||''); });
    return { ok:true, hoy:hoy, pendientes:pendientes, proximos:proximos,
             hechos:hechos.slice(0,40), totalPendientes:pendientes.length };
  } catch(ex){ return { ok:false, error:ex.message, pendientes:[], proximos:[], hechos:[] }; }
}

// Conteo ligero para la campana / banner.
function readRecordatoriosPendientes(usuario) {
  try { var r = readRecordatorios(usuario); return { ok:true, total: (r.pendientes||[]).length }; }
  catch(ex){ return { ok:false, total:0 }; }
}

// Crear — punto de entrada ABIERTO para cualquier área del sistema.
function saveRecordatorio(b) {
  try {
    var sh = _recSheet();
    var lr = sh.getLastRow(), max = 0;
    if (lr>1){ var ids=sh.getRange(2,1,lr-1,1).getValues();
      for (var i=0;i<ids.length;i++){ var m=String(ids[i][0]||'').match(/(\d+)/); if(m){var n=parseInt(m[1]); if(n>max)max=n;} } }
    var id = 'REC-'+String(max+1).padStart(5,'0');
    var fecha = b.fecha ? _recFmtDate(b.fecha) : _recToday();
    var row = [ id, String(b.usuario||'').toLowerCase().trim(), b.titulo||'Recordatorio', b.detalle||'',
                fecha, b.origen||'manual', b.referencia||'', 'pendiente', new Date(), '' ];
    sh.appendRow(row);
    return { ok:true, id:id, fecha:fecha };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

function updateRecordatorioEstado(b) {
  try {
    var sh=_recSheet(); var rn=b.rowNum; if(!rn||rn<2) return {ok:false, error:'Fila inválida'};
    var estado = (b.estado==='hecho'||b.estado==='descartado') ? b.estado : 'pendiente';
    sh.getRange(rn,8).setValue(estado);
    sh.getRange(rn,10).setValue(estado==='pendiente'?'':new Date());
    return {ok:true};
  } catch(ex){ return {ok:false, error:ex.message}; }
}

function deleteRecordatorio(b) {
  try { var sh=_recSheet(); var rn=b.rowNum; if(!rn||rn<2) return {ok:false,error:'Fila inválida'}; sh.deleteRow(rn); return {ok:true}; }
  catch(ex){ return {ok:false, error:ex.message}; }
}

// Tarea programada (aparece sola en Panel de Control → Programación):
// cada mañana manda a cada usuario un correo con sus pendientes del día.
function recordatoriosDelDia() {
  try {
    var rows=_recRows(), hoy=_recToday(), byUser={};
    rows.forEach(function(r){
      if (r.estado!=='pendiente') return;
      if (!(r.fecha && r.fecha<=hoy)) return;
      (byUser[r.usuario]=byUser[r.usuario]||[]).push(r);
    });
    Object.keys(byUser).forEach(function(email){
      if (!email) return;
      var lista=byUser[email].map(function(r){
        return '• '+r.titulo+(r.fecha<hoy?' (vencido '+r.fecha+')':'')+(r.referencia?'  ['+r.referencia+']':'');
      }).join('\n');
      try { MailApp.sendEmail(email, 'Tus recordatorios de hoy — '+hoy,
        'Tienes '+byUser[email].length+' recordatorio(s) para hoy:\n\n'+lista+'\n\nEntra al ERP → Recordatorios.'); } catch(e){}
    });
  } catch(e){ Logger.log('[recordatorios] recordatoriosDelDia: '+e.message); }
}
