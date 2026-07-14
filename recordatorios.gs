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
// Columnas 11-13 (Recurrencia/RecDia/RecMes) agregadas para recordatorios periódicos.
var REC_HEADERS = ['ID','Usuario','Titulo','Detalle','FechaObjetivo','Origen','Referencia','Estado','CreadoEn','CompletadoEn','Recurrencia','RecDia','RecMes'];
var _REC_DOW = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
var _REC_MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

function setupRecordatorios() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(REC_TAB);
  if (!sh) sh = ss.insertSheet(REC_TAB);
  sh.clear();
  sh.getRange(1,1,1,REC_HEADERS.length).setValues([REC_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, REC_HEADERS.length, [110,210,260,300,120,110,150,100,160,160,110,80,80]);
  return {ok:true, tab:REC_TAB};
}
// Asegura que la hoja tenga todas las columnas (migración de hojas viejas de 10 columnas).
function _recEnsureCols(sh) {
  try {
    var have = sh.getMaxColumns(), need = REC_HEADERS.length;
    if (have < need) sh.insertColumnsAfter(have, need-have);
    var cur = sh.getRange(1,1,1,need).getValues()[0];
    var falta = false;
    for (var i=0;i<need;i++){ if (String(cur[i]||'')!==REC_HEADERS[i]){ falta=true; break; } }
    if (falta) sh.getRange(1,1,1,need).setValues([REC_HEADERS]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  } catch(e){}
}
function _recSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(REC_TAB);
  if (!sh) { setupRecordatorios(); sh = ss.getSheetByName(REC_TAB); }
  _recEnsureCols(sh);
  return sh;
}
// ── Recurrencia ────────────────────────────────────────────────
function _recParseYmd(s){
  var m = String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(!m) return null;
  return new Date(parseInt(m[1],10), parseInt(m[2],10)-1, parseInt(m[3],10));
}
function _recYmd(d){
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
function _recDaysInMonth(y,mIndex){ return new Date(y, mIndex+1, 0).getDate(); }
// Próxima ocurrencia (YYYY-MM-DD) en o después de baseStr; si strictlyAfter, estrictamente después.
function _recNextOcc(rec, recDia, recMes, baseStr, strictlyAfter){
  rec = String(rec||'').toLowerCase();
  if(!rec || rec==='none' || rec==='ninguna') return baseStr;
  var base = _recParseYmd(baseStr) || new Date();
  base = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  recDia = parseInt(recDia,10); recMes = parseInt(recMes,10);
  var cand;
  if (rec==='diaria'){
    cand = new Date(base.getTime());
    if (strictlyAfter) cand.setDate(cand.getDate()+1);
    return _recYmd(cand);
  }
  if (rec==='semanal'){
    var dow = isNaN(recDia)?base.getDay():recDia;
    cand = new Date(base.getTime());
    var add = ((dow - cand.getDay()) + 7) % 7;
    if (add===0 && strictlyAfter) add = 7;
    cand.setDate(cand.getDate()+add);
    return _recYmd(cand);
  }
  if (rec==='mensual'){
    var d = isNaN(recDia)?base.getDate():recDia;
    var y=base.getFullYear(), mo=base.getMonth();
    for (var k=0;k<420;k++){
      var dim=_recDaysInMonth(y,mo), dd=Math.min(d,dim);
      cand=new Date(y,mo,dd);
      var ok = strictlyAfter ? (cand.getTime()>base.getTime()) : (cand.getTime()>=base.getTime());
      if (ok) return _recYmd(cand);
      mo++; if(mo>11){mo=0;y++;}
    }
  }
  if (rec==='anual'){
    var dA=isNaN(recDia)?base.getDate():recDia, mA=isNaN(recMes)?base.getMonth():(recMes-1);
    var yy=base.getFullYear();
    for (var j=0;j<80;j++){
      var dimA=_recDaysInMonth(yy,mA), ddA=Math.min(dA,dimA);
      cand=new Date(yy,mA,ddA);
      var okA = strictlyAfter ? (cand.getTime()>base.getTime()) : (cand.getTime()>=base.getTime());
      if (okA) return _recYmd(cand);
      yy++;
    }
  }
  return baseStr;
}
function _recRecurLabel(rec, recDia, recMes){
  rec=String(rec||'').toLowerCase();
  if(rec==='diaria') return 'Cada día';
  if(rec==='semanal') return 'Cada '+(_REC_DOW[parseInt(recDia,10)]||'semana');
  if(rec==='mensual') return 'Cada mes · día '+(parseInt(recDia,10)||1);
  if(rec==='anual') return 'Cada año · '+(parseInt(recDia,10)||1)+' de '+(_REC_MESES[(parseInt(recMes,10)||1)-1]||'');
  return '';
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
    var rec = String(r[10]||'').toLowerCase();
    rows.push({
      rowNum:i+1, id:String(r[0]), usuario:String(r[1]||'').toLowerCase().trim(),
      titulo:String(r[2]||''), detalle:String(r[3]||''), fecha:_recFmtDate(r[4]),
      origen:String(r[5]||'manual'), referencia:String(r[6]||''),
      estado:String(r[7]||'pendiente').toLowerCase(), creadoEn:_recFmtDate(r[8]),
      completadoEn:_recFmtDate(r[9]),
      recurrencia:rec, recDia:(r[11]===''||r[11]===undefined)?'':parseInt(r[11],10),
      recMes:(r[12]===''||r[12]===undefined)?'':parseInt(r[12],10),
      recLabel:_recRecurLabel(rec, r[11], r[12])
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
    var rec = String(b.recurrencia||'').toLowerCase(); if(rec==='none'||rec==='ninguna') rec='';
    var recDia = (rec && b.recDia!==undefined && b.recDia!=='') ? parseInt(b.recDia,10) : '';
    var recMes = (rec && b.recMes!==undefined && b.recMes!=='') ? parseInt(b.recMes,10) : '';
    var desde = b.fecha ? _recFmtDate(b.fecha) : _recToday();
    var fecha = rec ? _recNextOcc(rec, recDia, recMes, desde, false) : desde;
    var row = [ id, String(b.usuario||'').toLowerCase().trim(), b.titulo||'Recordatorio', b.detalle||'',
                fecha, b.origen||'manual', b.referencia||'', 'pendiente', new Date(), '',
                rec, recDia, recMes ];
    sh.appendRow(row);
    return { ok:true, id:id, fecha:fecha, recurrente:!!rec };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

function updateRecordatorioEstado(b) {
  try {
    var sh=_recSheet(); var rn=b.rowNum; if(!rn||rn<2) return {ok:false, error:'Fila inválida'};
    var estado = (b.estado==='hecho'||b.estado==='descartado') ? b.estado : 'pendiente';
    // Un recordatorio recurrente marcado "hecho" NO se archiva: avanza al próximo periodo.
    if (estado==='hecho') {
      var rv = sh.getRange(rn,1,1,REC_HEADERS.length).getValues()[0];
      var rec = String(rv[10]||'').toLowerCase();
      if (rec && rec!=='none' && rec!=='ninguna') {
        var prox = _recNextOcc(rec, rv[11], rv[12], _recToday(), true);
        sh.getRange(rn,5).setValue(prox);        // nueva fecha objetivo
        sh.getRange(rn,8).setValue('pendiente'); // sigue vivo para el próximo periodo
        sh.getRange(rn,10).setValue(new Date()); // última vez hecho
        return {ok:true, recurrente:true, proxima:prox};
      }
    }
    sh.getRange(rn,8).setValue(estado);
    sh.getRange(rn,10).setValue(estado==='pendiente'?'':new Date());
    return {ok:true};
  } catch(ex){ return {ok:false, error:ex.message}; }
}

function deleteRecordatorio(b) {
  try { var sh=_recSheet(); var rn=b.rowNum; if(!rn||rn<2) return {ok:false,error:'Fila inválida'}; sh.deleteRow(rn); return {ok:true}; }
  catch(ex){ return {ok:false, error:ex.message}; }
}

// Editar / reprogramar: cambia título, detalle y/o fecha (col 3, 4, 5).
// Si b.reabrir, lo vuelve a poner pendiente (útil para reprogramar uno ya hecho).
function updateRecordatorio(b) {
  try {
    var sh=_recSheet(); var rn=b.rowNum; if(!rn||rn<2) return {ok:false, error:'Fila inválida'};
    if(b.titulo!==undefined)  sh.getRange(rn,3).setValue(b.titulo||'Recordatorio');
    if(b.detalle!==undefined) sh.getRange(rn,4).setValue(b.detalle||'');
    if(b.recurrencia!==undefined){
      var rec2=String(b.recurrencia||'').toLowerCase(); if(rec2==='none'||rec2==='ninguna') rec2='';
      var rd2=(rec2 && b.recDia!==undefined && b.recDia!=='')?parseInt(b.recDia,10):'';
      var rm2=(rec2 && b.recMes!==undefined && b.recMes!=='')?parseInt(b.recMes,10):'';
      sh.getRange(rn,11).setValue(rec2);
      sh.getRange(rn,12).setValue(rd2);
      sh.getRange(rn,13).setValue(rm2);
      if(rec2 && !b.fecha) sh.getRange(rn,5).setValue(_recNextOcc(rec2, rd2, rm2, _recToday(), false));
    }
    if(b.fecha){
      var rec3=String(b.recurrencia!==undefined?b.recurrencia:'').toLowerCase();
      sh.getRange(rn,5).setValue(rec3 && rec3!=='none' && rec3!=='ninguna'
        ? _recNextOcc(rec3, b.recDia, b.recMes, _recFmtDate(b.fecha), false)
        : _recFmtDate(b.fecha));
    }
    if(b.reabrir){ sh.getRange(rn,8).setValue('pendiente'); sh.getRange(rn,10).setValue(''); }
    return {ok:true};
  } catch(ex){ return {ok:false, error:ex.message}; }
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
