/* ══════════════════════════════════════════════════════════════
   GASTOS FIJOS — plantillas recurrentes + programación mensual a CxP
   Hoja "GastosFijos" en el spreadsheet de Egresos (EGRESOS_SS_2026).
   Las instancias se crean como filas CxP en Egresos2026, marcadas con
   la columna RecurrenteID (autocreada) para evitar duplicados y
   prellenar el siguiente mes con el monto real anterior.
   ══════════════════════════════════════════════════════════════ */

var GF_TAB = 'GastosFijos';
// Columnas: A ID, B Activo, C Proveedor, D Contable, E Subtipo, F Concepto,
//           G MontoEstimado, H MontoVariable, I DiaVencimiento, J Meses,
//           K Desde, L Hasta, M FormaPago, N Notas
var GF_HEADERS = ['ID','Activo','Proveedor','Contable','Subtipo','Concepto',
  'MontoEstimado','MontoVariable','DiaVencimiento','Meses','Desde','Hasta','FormaPago','Notas'];

function setupGastosFijos() {
  var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
  var sh = ss.getSheetByName(GF_TAB);
  if (!sh) sh = ss.insertSheet(GF_TAB);
  sh.clear();
  sh.getRange(1,1,1,GF_HEADERS.length).setValues([GF_HEADERS]).setFontWeight('bold');
  // Ejemplos
  var ej = [
    ['GF-001', true, 'Arrendador',  'Gasto', 'Renta',    'Renta laboratorio', 41067.08, false, '5',   'Todos', '', '', 'Santander', ''],
    ['GF-002', true, 'Nómina',      'Gasto', 'Nomina',   'Nómina 1ra quincena', 0,       true,  '15',  'Todos', '', '', 'Santander', 'Varía por bonos'],
    ['GF-003', true, 'Nómina',      'Gasto', 'Nomina',   'Nómina 2da quincena', 0,       true,  'fin', 'Todos', '', '', 'Santander', ''],
    ['GF-004', true, 'LIFEAIRE',    'Gasto', 'Mantenimiento', 'Servicio LIFEAIRE', 0,    false, '10',  'Todos', '', '', 'Santander', '']
  ];
  sh.getRange(2,1,ej.length,GF_HEADERS.length).setValues(ej);
  sh.setFrozenRows(1);
  return {ok:true, tab:GF_TAB};
}

function _gfSheet() {
  var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
  var sh = ss.getSheetByName(GF_TAB);
  if (!sh) { setupGastosFijos(); sh = ss.getSheetByName(GF_TAB); }
  return sh;
}

function readGastosFijos() {
  try {
    var sh = _gfSheet();
    var v = sh.getDataRange().getValues();
    var rows = [];
    for (var i=1;i<v.length;i++) {
      var t=v[i]; var id=String(t[0]||'').trim(); if(!id) continue;
      rows.push({
        rowNum:i+1, id:id,
        activo: t[1]===true||String(t[1]).toUpperCase()==='TRUE',
        proveedor:String(t[2]||''), contable:String(t[3]||'Gasto'), subtipo:String(t[4]||''),
        concepto:String(t[5]||''), montoEstimado:parseFloat(t[6])||0,
        montoVariable: t[7]===true||String(t[7]).toUpperCase()==='TRUE',
        diaVencimiento:String(t[8]||''), meses:String(t[9]||'Todos'),
        desde:String(t[10]||''), hasta:String(t[11]||''),
        formaPago:String(t[12]||''), notas:String(t[13]||'')
      });
    }
    return {ok:true, rows:rows};
  } catch(ex) { return {ok:false, error:ex.message, rows:[]}; }
}

function _gfRowFromBody(b) {
  return [
    b.id||'', b.activo===true||String(b.activo).toUpperCase()==='TRUE',
    b.proveedor||'', b.contable||'Gasto', b.subtipo||'', b.concepto||'',
    parseFloat(String(b.montoEstimado||'').replace(/[$,]/g,''))||0,
    b.montoVariable===true||String(b.montoVariable).toUpperCase()==='TRUE',
    String(b.diaVencimiento||''), String(b.meses||'Todos'),
    String(b.desde||''), String(b.hasta||''), b.formaPago||'', b.notas||''
  ];
}

function saveGastoFijo(b) {
  try {
    var sh=_gfSheet();
    // ID consecutivo GF-00X
    var lr=sh.getLastRow(), max=0;
    if(lr>1){ var ids=sh.getRange(2,1,lr-1,1).getValues();
      for(var i=0;i<ids.length;i++){ var m=String(ids[i][0]||'').match(/(\d+)/); if(m){var n=parseInt(m[1]); if(n>max)max=n;} } }
    b.id = 'GF-'+String(max+1).padStart(3,'0');
    if(b.activo===undefined) b.activo=true;
    sh.appendRow(_gfRowFromBody(b));
    logAudit(b.usuario||'sistema','GastoFijo','Crear',b.id,'',' ',b.proveedor+' | '+b.concepto);
    return {ok:true, id:b.id};
  } catch(ex){ return {ok:false, error:ex.message}; }
}

function updateGastoFijo(b) {
  try {
    var sh=_gfSheet();
    var rn=b.rowNum; if(!rn||rn<2) return {ok:false, error:'Fila inválida'};
    var cur=sh.getRange(rn,1).getValue(); // conservar ID
    b.id = b.id || cur;
    sh.getRange(rn,1,1,GF_HEADERS.length).setValues([_gfRowFromBody(b)]);
    logAudit(b.usuario||'sistema','GastoFijo','Editar',b.id,'',' ',b.proveedor||'');
    return {ok:true, id:b.id};
  } catch(ex){ return {ok:false, error:ex.message}; }
}

function toggleGastoFijo(b) {
  try {
    var sh=_gfSheet(); var rn=b.rowNum; if(!rn||rn<2) return {ok:false, error:'Fila inválida'};
    sh.getRange(rn,2).setValue(b.activo===true||String(b.activo).toUpperCase()==='TRUE');
    return {ok:true};
  } catch(ex){ return {ok:false, error:ex.message}; }
}

function deleteGastoFijo(b) {
  try {
    var sh=_gfSheet(); var rn=b.rowNum; if(!rn||rn<2) return {ok:false, error:'Fila inválida'};
    sh.deleteRow(rn);
    return {ok:true};
  } catch(ex){ return {ok:false, error:ex.message}; }
}

function _gfVencimiento(periodo, dia) {
  var y=parseInt(periodo.substring(0,4),10), m=parseInt(periodo.substring(5,7),10);
  var lastDay=new Date(y, m, 0).getDate();
  var d;
  if(!dia || String(dia).toLowerCase()==='fin') d=lastDay;
  else { d=parseInt(dia,10)||lastDay; if(d>lastDay)d=lastDay; if(d<1)d=1; }
  return periodo+'-'+String(d).padStart(2,'0');
}

// Lee Egresos2026 una vez: devuelve {iRec, genSet, lastByGF}
function _gfEgContext() {
  var ss=SpreadsheetApp.openById(EGRESOS_SS_2026);
  var egSh=ss.getSheetByName(EGRESOS_TABS[2026]||'Egresos2026');
  var egHdr=egSh.getRange(1,1,1,egSh.getLastColumn()).getValues()[0].map(function(h){return String(h).trim().toLowerCase();});
  var iRec=-1; for(var c=0;c<egHdr.length;c++){ if(egHdr[c].indexOf('recurrente')>-1){iRec=c;break;} } // 0-indexed
  var genSet={}, lastByGF={};
  if(iRec>-1){
    var d=egSh.getDataRange().getValues();
    for(var i=1;i<d.length;i++){
      var rec=String(d[i][iRec]||'').trim(); if(!rec) continue;
      var mes=String(d[i][2]||'').trim();
      genSet[rec+'|'+mes]=true;
      var monto=parseFloat(d[i][9])||0;
      if(!lastByGF[rec] || mes>lastByGF[rec].mes) lastByGF[rec]={mes:mes, monto:monto};
    }
  }
  return {egSh:egSh, iRec:iRec, genSet:genSet, lastByGF:lastByGF};
}

function _gfAplica(t, periodo, mesNum) {
  var activo = t[1]===true||String(t[1]).toUpperCase()==='TRUE';
  if(!activo) return false;
  var desde=String(t[10]||'').trim(), hasta=String(t[11]||'').trim();
  if(desde && periodo<desde) return false;
  if(hasta && periodo>hasta) return false;
  var meses=String(t[9]||'Todos').trim();
  if(meses && meses.toLowerCase()!=='todos'){
    var list=meses.split(/[,\s]+/).map(function(x){return parseInt(x,10);}).filter(function(x){return x>=1&&x<=12;});
    if(list.indexOf(mesNum)===-1) return false;
  }
  return true;
}

function _gfItem(t, periodo, lastByGF) {
  var id=String(t[0]||'').trim();
  var estimado=parseFloat(t[6])||0;
  var sugerido=(lastByGF[id]&&lastByGF[id].monto)?lastByGF[id].monto:estimado;
  return {
    id:id, proveedor:String(t[2]||''), contable:String(t[3]||'Gasto'),
    subtipo:String(t[4]||''), concepto:String(t[5]||''),
    montoEstimado:estimado, montoSugerido:sugerido,
    montoVariable: t[7]===true||String(t[7]).toUpperCase()==='TRUE',
    diaVencimiento:String(t[8]||''), vencimiento:_gfVencimiento(periodo, String(t[8]||'')),
    formaPago:String(t[12]||''), notas:String(t[13]||''),
    ultimoReal: lastByGF[id]||null
  };
}

// Propuestas del mes: plantillas aplicables aún no generadas (+ las ya generadas).
function readGastosFijosPropuestas(periodo) {
  try {
    if(!periodo){ var d=new Date(); periodo=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'); }
    var mesNum=parseInt(periodo.substring(5,7),10);
    var sh=_gfSheet(); var gf=sh.getDataRange().getValues();
    var ctx=_gfEgContext();
    var propuestas=[], generadas=[];
    for(var j=1;j<gf.length;j++){
      var t=gf[j]; if(!String(t[0]||'').trim()) continue;
      if(!_gfAplica(t, periodo, mesNum)) continue;
      var item=_gfItem(t, periodo, ctx.lastByGF);
      if(ctx.genSet[item.id+'|'+periodo]) generadas.push(item); else propuestas.push(item);
    }
    return {ok:true, periodo:periodo, propuestas:propuestas, generadas:generadas,
            pendientes:propuestas.length};
  } catch(ex){ return {ok:false, error:ex.message, propuestas:[], generadas:[]}; }
}

// Proyección a N meses (default 3 = trimestre) — SOLO lectura, para budget. No crea nada.
function readProyeccionGastosFijos(b) {
  try {
    var n=parseInt(b&&b.meses)||3; if(n<1)n=1; if(n>12)n=12;
    var base=(b&&b.periodo)||(function(){var d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');})();
    var y=parseInt(base.substring(0,4),10), m=parseInt(base.substring(5,7),10)-1;
    var sh=_gfSheet(); var gf=sh.getDataRange().getValues();
    var ctx=_gfEgContext();
    var periodos=[];
    for(var k=0;k<n;k++){
      var dd=new Date(y, m+k, 1);
      var per=dd.getFullYear()+'-'+String(dd.getMonth()+1).padStart(2,'0');
      var mesNum=dd.getMonth()+1;
      var items=[], total=0, porCat={};
      for(var j=1;j<gf.length;j++){
        var t=gf[j]; if(!String(t[0]||'').trim()) continue;
        if(!_gfAplica(t, per, mesNum)) continue;
        var it=_gfItem(t, per, ctx.lastByGF);
        items.push(it); total+=it.montoSugerido;
        var c=it.subtipo||'—'; porCat[c]=(porCat[c]||0)+it.montoSugerido;
      }
      periodos.push({periodo:per, total:total, items:items, porCategoria:porCat});
    }
    return {ok:true, periodos:periodos};
  } catch(ex){ return {ok:false, error:ex.message, periodos:[]}; }
}

// Crea UNA fila CxP en Egresos2026 a partir de un item validado.
function _gfAppendCxP(egSh, iRec1, item, usuario) {
  var periodo=item.periodo;
  var monto=parseFloat(String(item.monto||'').replace(/[$,]/g,''))||0;
  var data=egSh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(String(data[i][iRec1-1]||'').trim()===item.id && String(data[i][2]||'').trim()===periodo)
      return {ok:false, dup:true, id:item.id};
  }
  var lr=egSh.getLastRow(), lastId=0;
  if(lr>1){ var ids=egSh.getRange(2,1,lr-1,1).getValues(); for(var k=0;k<ids.length;k++){var n=parseInt(ids[k][0]); if(n>lastId)lastId=n;} }
  var newId=lastId+1;
  // A ID, B Fecha(''), C Mes, D prioridad, E Proveedor, F Contable, G Tipo, H Subtipo,
  // I Concepto, J Monto, K Notas, L Vencimiento, M..O false, P Poliza, Q FormaPago, R Obs, S/T links
  var row=[ newId,'',periodo,1, item.proveedor||'', item.contable||'Gasto','Fijo', item.subtipo||'',
            item.concepto||'', monto, item.notas||'', item.vencimiento||'', false,false,false,'',
            item.formaPago||'', '', '', '' ];
  egSh.appendRow(row);
  egSh.getRange(egSh.getLastRow(), iRec1).setValue(item.id); // RecurrenteID
  logAudit(usuario||'sistema','GastoFijo','Programar',item.id,'Mes',periodo,(item.proveedor||'')+' $'+monto);
  return {ok:true, id:newId};
}

function programarGastoFijo(b) {
  try {
    var ss=SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egSh=ss.getSheetByName(EGRESOS_TABS[2026]||'Egresos2026');
    var iRec1=_egColEnsure(egSh,'recurrente','RecurrenteID'); // 1-indexed
    var res=_gfAppendCxP(egSh, iRec1, b, b.usuario);
    return res;
  } catch(ex){ return {ok:false, error:ex.message}; }
}

function programarGastosFijosBatch(b) {
  try {
    var items=b.items||[]; if(!items.length) return {ok:false, error:'Sin partidas'};
    var ss=SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egSh=ss.getSheetByName(EGRESOS_TABS[2026]||'Egresos2026');
    var iRec1=_egColEnsure(egSh,'recurrente','RecurrenteID');
    var creadas=0, dups=0, errores=0;
    for(var i=0;i<items.length;i++){
      items[i].usuario=b.usuario;
      var r=_gfAppendCxP(egSh, iRec1, items[i], b.usuario);
      if(r.ok) creadas++; else if(r.dup) dups++; else errores++;
    }
    return {ok:true, creadas:creadas, duplicadas:dups, errores:errores};
  } catch(ex){ return {ok:false, error:ex.message}; }
}
