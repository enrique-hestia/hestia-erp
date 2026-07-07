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
//           K Desde, L Hasta, M FormaPago, N Notas, O Divisa, P Frecuencia
var GF_HEADERS = ['ID','Activo','Proveedor','Contable','Subtipo','Concepto',
  'MontoEstimado','MontoVariable','DiaVencimiento','Meses','Desde','Hasta','FormaPago','Notas','Divisa','Frecuencia'];

function setupGastosFijos() {
  var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
  var sh = ss.getSheetByName(GF_TAB);
  var isNew = !sh;
  if (isNew) sh = ss.insertSheet(GF_TAB);

  if (isNew || sh.getLastRow() < 2) {
    // Solo inicializa si la hoja es nueva o está vacía — NUNCA borra datos existentes
    if (isNew) sh.getRange(1,1,1,GF_HEADERS.length).setValues([GF_HEADERS]).setFontWeight('bold');
    var ej = [
      ['GF-001', true, 'Arrendador',  'Gasto', 'Renta',    'Renta laboratorio', 41067.08, false, '5',   'Todos', '', '', 'Santander', '', 'MXN', 'mensual'],
      ['GF-002', true, 'Nómina',      'Gasto', 'Nomina',   'Nómina quincenal',   0,       true,  '15',  'Todos', '', '', 'Santander', 'Varía por bonos', 'MXN', 'quincena-ambas'],
      ['GF-004', true, 'LIFEAIRE',    'Gasto', 'Mantenimiento', 'Servicio LIFEAIRE', 0,  false, '10',  'Todos', '', '', 'AMEX', 'Pago en dólares', 'USD', 'mensual']
    ];
    sh.getRange(2,1,ej.length,GF_HEADERS.length).setValues(ej);
    sh.setFrozenRows(1);
    return {ok:true, tab:GF_TAB, created:true};
  }

  // La hoja ya tiene datos: solo asegura los headers nuevos sin tocar nada más
  try {
    var lc = sh.getLastColumn();
    if (lc < GF_HEADERS.length) {
      sh.getRange(1, lc+1, 1, GF_HEADERS.length-lc).setValues([GF_HEADERS.slice(lc)]);
    }
  } catch(e){}
  return {ok:true, tab:GF_TAB, created:false, note:'Hoja existente — solo se actualizaron headers'};
}

function _gfSheet() {
  var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
  var sh = ss.getSheetByName(GF_TAB);
  if (!sh) { setupGastosFijos(); sh = ss.getSheetByName(GF_TAB); }
  // Asegura headers nuevos (Divisa col O, Frecuencia col P) en hojas creadas antes — sin borrar datos.
  try {
    var lc = sh.getLastColumn();
    if (lc < GF_HEADERS.length) {
      sh.getRange(1, lc+1, 1, GF_HEADERS.length-lc).setValues([GF_HEADERS.slice(lc)]);
    }
  } catch(e){}
  return sh;
}

function readGastosFijos() {
  try {
    var sh = _gfSheet();
    var v = sh.getDataRange().getValues();
    var ctx = null;
    try { ctx = _gfEgContext(); } catch(e) {}
    var rows = [];
    for (var i=1;i<v.length;i++) {
      var t=v[i]; var id=String(t[0]||'').trim(); if(!id) continue;
      var last = ctx ? (ctx.lastByGF[id]||ctx.lastByGF[id+'_Q1']||ctx.lastByGF[id+'_Q2']||null) : null;
      rows.push({
        rowNum:i+1, id:id,
        activo: t[1]===true||String(t[1]).toUpperCase()==='TRUE',
        proveedor:String(t[2]||''), contable:String(t[3]||'Gasto'), subtipo:String(t[4]||''),
        concepto:String(t[5]||''), montoEstimado:parseFloat(t[6])||0,
        montoVariable: t[7]===true||String(t[7]).toUpperCase()==='TRUE',
        diaVencimiento:String(t[8]||''), meses:String(t[9]||'Todos'),
        desde:String(t[10]||''), hasta:String(t[11]||''),
        formaPago:String(t[12]||''), notas:String(t[13]||''),
        divisa:String(t[14]||'MXN').toUpperCase()==='USD'?'USD':'MXN',
        frecuencia:String(t[15]||'mensual'),
        lastProgramado: last ? last.mes : ''
      });
    }
    return {ok:true, rows:rows};
  } catch(ex) { return {ok:false, error:ex.message, rows:[]}; }
}

// Reconstruye el catálogo GastosFijos leyendo los RecurrenteID únicos en Egresos2026.
// Solo crea entradas que aún no existen en GastosFijos — nunca sobreescribe.
function reconstruirCatalogoGF(b) {
  try {
    var ss  = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egSh = ss.getSheetByName(EGRESOS_TABS[2026]||'Egresos2026');
    var gfSh = _gfSheet();

    // Localizar columnas dinámicas en Egresos
    var egHdr = egSh.getRange(1,1,1,egSh.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).trim().toLowerCase(); });
    var iRec=-1, iDiv=-1;
    for(var c=0;c<egHdr.length;c++){
      if(egHdr[c].indexOf('recurrente')>-1) iRec=c;
      if(egHdr[c]==='divisa') iDiv=c;
    }
    if(iRec<0) return {ok:false,error:'No se encontró columna RecurrenteID en Egresos'};

    // Leer y agrupar filas de Egresos por base-ID de RecurrenteID
    var data = egSh.getDataRange().getValues();
    var groups = {};
    for(var i=1;i<data.length;i++){
      var rec=String(data[i][iRec]||'').trim(); if(!rec) continue;
      var base=rec.replace(/_Q[12]$/,'');
      var isQ1=/_Q1$/.test(rec), isQ2=/_Q2$/.test(rec);
      var mes=String(data[i][2]||'').trim();
      if(!groups[base]) groups[base]={id:base,lastMes:'',lastRow:null,hasQ1:false,hasQ2:false};
      var g=groups[base];
      if(isQ1) g.hasQ1=true;
      if(isQ2) g.hasQ2=true;
      if(!g.lastMes||mes>g.lastMes){ g.lastMes=mes; g.lastRow=data[i]; }
    }

    // IDs que ya existen en GastosFijos
    var existing={};
    var gfData=gfSh.getDataRange().getValues();
    var maxNum=0;
    for(var j=1;j<gfData.length;j++){
      var gid=String(gfData[j][0]||'').trim(); if(!gid) continue;
      existing[gid]=true;
      var mn=gid.match(/(\d+)/); if(mn){var n=parseInt(mn[1]);if(n>maxNum)maxNum=n;}
    }

    var created=0;
    var sortedIds=Object.keys(groups).sort();
    for(var k=0;k<sortedIds.length;k++){
      var bid=sortedIds[k];
      if(existing[bid]) continue; // Ya existe — no tocar

      var g=groups[bid];
      var row=g.lastRow;
      var freq=g.hasQ1&&g.hasQ2?'quincena-ambas':g.hasQ1?'quincena-15':g.hasQ2?'quincena-30':'mensual';

      // Inferir día de vencimiento del último registro
      var venc=String(row[11]||'').trim(), diaVenc='';
      if(venc&&venc.length>=10){
        var dd=parseInt(venc.substring(8,10),10);
        var lastDay=new Date(parseInt(venc.substring(0,4)),parseInt(venc.substring(5,7)),0).getDate();
        diaVenc=(dd===lastDay)?'fin':String(dd);
      }

      var div=(iDiv>-1&&row[iDiv])?String(row[iDiv]).toUpperCase():'MXN';
      var monto=parseFloat(row[9])||0;
      // Limpiar sufijos de quincena del concepto
      var concepto=String(row[8]||'').replace(/\s*\([12]ª quincena\)/gi,'').trim();

      // Usar el ID original del RecurrenteID; si no tiene formato GF-NNN, asignar nuevo
      var newId=bid.match(/^GF-\d+$/)?bid:('GF-'+String(++maxNum).padStart(3,'0'));

      var gfRow=[
        newId, true,
        String(row[4]||''),         // Proveedor
        String(row[5]||'Gasto'),    // Contable
        String(row[7]||''),         // Subtipo
        concepto,                    // Concepto
        monto, false,                // MontoEstimado, MontoVariable
        diaVenc, 'Todos', '', '',    // DiaVencimiento, Meses, Desde, Hasta
        String(row[16]||''),         // FormaPago (col Q)
        String(row[10]||''),         // Notas
        div==='USD'?'USD':'MXN',     // Divisa
        freq                          // Frecuencia
      ];
      gfSh.appendRow(gfRow);
      existing[newId]=true;
      created++;
    }
    logAudit(b&&b.usuario||'sistema','GastoFijo','Reconstruir','','','',created+' entradas desde BD_CxP');
    return {ok:true,created:created,omitidos:sortedIds.length-created};
  } catch(ex){ return {ok:false,error:ex.message}; }
}

function _gfRowFromBody(b) {
  return [
    b.id||'', b.activo===true||String(b.activo).toUpperCase()==='TRUE',
    b.proveedor||'', b.contable||'Gasto', b.subtipo||'', b.concepto||'',
    parseFloat(String(b.montoEstimado||'').replace(/[$,]/g,''))||0,
    b.montoVariable===true||String(b.montoVariable).toUpperCase()==='TRUE',
    String(b.diaVencimiento||''), String(b.meses||'Todos'),
    String(b.desde||''), String(b.hasta||''), b.formaPago||'', b.notas||'',
    String(b.divisa||'MXN').toUpperCase()==='USD'?'USD':'MXN',
    String(b.frecuencia||'mensual')
  ];
}

function saveGastoFijo(b) {
  try {
    var sh=_gfSheet();
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

function saveGastosFijosBatch(b) {
  try {
    var items=b.items||[]; if(!items.length) return {ok:false,error:'Sin partidas'};
    var sh=_gfSheet();
    var lr=sh.getLastRow(), max=0;
    if(lr>1){ var ids=sh.getRange(2,1,lr-1,1).getValues();
      for(var i=0;i<ids.length;i++){ var m=String(ids[i][0]||'').match(/(\d+)/); if(m){var n=parseInt(m[1]);if(n>max)max=n;} } }
    var saved=0, errors=[];
    for(var j=0;j<items.length;j++){
      try{
        var it=items[j]; it.activo=true; max++;
        it.id='GF-'+String(max).padStart(3,'0');
        sh.appendRow(_gfRowFromBody(it));
        saved++;
      }catch(ex){ errors.push(ex.message); }
    }
    logAudit(b.usuario||'sistema','GastoFijo','CrearBatch','','','',saved+' partidas');
    return {ok:true,saved:saved,errors:errors};
  }catch(ex){ return {ok:false,error:ex.message}; }
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
      // OJO: la columna Mes a veces la guarda Sheets como objeto Date en vez de
      // texto "YYYY-MM" (según cómo se haya capturado la fila) — sin normalizar,
      // la comparación de "ya generado" fallaba en silencio y el gasto fijo
      // volvía a aparecer disponible para programar, aunque ya tuviera CxP.
      var mes=_gfToYM(d[i][2]);
      genSet[rec+'|'+mes]=true;
      var monto=parseFloat(d[i][9])||0;
      if(!lastByGF[rec] || mes>lastByGF[rec].mes) lastByGF[rec]={mes:mes, monto:monto};
      // Also index by base ID (strip _Q1/_Q2) so quincena items can find their historical monto
      var baseRec = rec.replace(/_Q[12]$/, '');
      if(baseRec !== rec && (!lastByGF[baseRec] || mes>lastByGF[baseRec].mes))
        lastByGF[baseRec]={mes:mes, monto:monto};
    }
  }
  return {egSh:egSh, iRec:iRec, genSet:genSet, lastByGF:lastByGF};
}

// Convierte cualquier valor de fecha a string "YYYY-MM" para comparación.
// Handles: Date objects de Sheets, strings "YYYY-MM", strings "YYYY-MM-DD",
// strings como "Wed Jul 01 2026 00:00:00 GMT+0000", y celdas vacías.
function _gfToYM(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0');
  }
  var s = String(v).trim();
  if (!s) return '';
  // Ya está en YYYY-MM o YYYY-MM-DD
  if (/^\d{4}-\d{2}/.test(s)) return s.substring(0,7);
  // Cualquier otro formato de fecha → parsear
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
  return s;
}

function _gfAplica(t, periodo, mesNum) {
  var activo = t[1]===true||String(t[1]).toUpperCase()==='TRUE';
  if(!activo) return false;
  var desde = _gfToYM(t[10]);
  var hasta  = _gfToYM(t[11]);
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
  var freq=String(t[15]||'mensual').trim();
  var estimado=parseFloat(t[6])||0;
  var last=lastByGF[id]||lastByGF[id+'_Q1']||lastByGF[id+'_Q2']||null;
  var sugerido=(last&&last.monto)?last.monto:estimado;
  return {
    id:id, proveedor:String(t[2]||''), contable:String(t[3]||'Gasto'),
    subtipo:String(t[4]||''), concepto:String(t[5]||''),
    montoEstimado:estimado, montoSugerido:sugerido,
    montoVariable: t[7]===true||String(t[7]).toUpperCase()==='TRUE',
    diaVencimiento:String(t[8]||''), vencimiento:_gfVencimiento(periodo, String(t[8]||'')),
    formaPago:String(t[12]||''), notas:String(t[13]||''),
    divisa:String(t[14]||'MXN').toUpperCase()==='USD'?'USD':'MXN',
    frecuencia:freq,
    ultimoReal: last
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
      var freq=item.frecuencia||'mensual';
      var isGen;
      if(freq==='quincena-ambas') isGen=!!ctx.genSet[item.id+'_Q1|'+periodo]&&!!ctx.genSet[item.id+'_Q2|'+periodo];
      else if(freq==='quincena-15') isGen=!!ctx.genSet[item.id+'_Q1|'+periodo];
      else if(freq==='quincena-30') isGen=!!ctx.genSet[item.id+'_Q2|'+periodo];
      else isGen=!!ctx.genSet[item.id+'|'+periodo];
      if(isGen) generadas.push(item); else propuestas.push(item);
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
// Usa LockService para que la revisión de duplicados + el appendRow sean atómicos:
// sin esto, dos clics casi simultáneos (doble clic, doble pestaña) podían pasar
// ambos la revisión de duplicado antes de que cualquiera terminara de escribir,
// generando dos cuentas por pagar del mismo gasto fijo y riesgo de pago doble.
function _gfAppendCxP(egSh, iRec1, item, usuario) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return {ok:false, error:'No se pudo obtener el bloqueo, intenta de nuevo'};
  try {
    var periodo=item.periodo;
    var monto=parseFloat(String(item.monto||'').replace(/[$,]/g,''))||0;
    var data=egSh.getDataRange().getValues();
    for(var i=1;i<data.length;i++){
      if(String(data[i][iRec1-1]||'').trim()===item.id && _gfToYM(data[i][2])===periodo)
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
    var newRow = egSh.getLastRow();
    egSh.getRange(newRow, iRec1).setValue(item.id); // RecurrenteID
    // Divisa de la partida (columna detectada/creada por header — no colisiona con otras)
    var div = String(item.divisa||'MXN').toUpperCase()==='USD' ? 'USD' : 'MXN';
    var iDiv = _egColEnsure(egSh, 'divisa', 'Divisa');
    egSh.getRange(newRow, iDiv).setValue(div);
    logAudit(usuario||'sistema','GastoFijo','Programar',item.id,'Mes',periodo,(item.proveedor||'')+' '+div+' '+monto);
  } finally {
    lock.releaseLock();
  }
  return {ok:true, id:newId};
}

// Crea las filas CxP de un gasto fijo respetando su frecuencia.
function _gfProgramarConFrec(egSh, iRec1, b, usuario) {
  var freq = String(b.frecuencia||'mensual').toLowerCase().trim();
  var creadas=0, dups=0;
  function clone(extra){ return Object.assign(JSON.parse(JSON.stringify(b)), extra); }
  function run(item){ var r=_gfAppendCxP(egSh,iRec1,item,usuario); if(r.ok)creadas++; else if(r.dup)dups++; return r; }
  if(freq==='quincena-ambas'){
    run(clone({id:b.id+'_Q1', vencimiento:_gfVencimiento(b.periodo,'15'),  concepto:(b.concepto||'')+' (1ª quincena)'}));
    run(clone({id:b.id+'_Q2', vencimiento:_gfVencimiento(b.periodo,'fin'), concepto:(b.concepto||'')+' (2ª quincena)'}));
  } else if(freq==='quincena-15'){
    run(clone({id:b.id+'_Q1', vencimiento:_gfVencimiento(b.periodo,'15'),  concepto:(b.concepto||'')+' (1ª quincena)'}));
  } else if(freq==='quincena-30'){
    run(clone({id:b.id+'_Q2', vencimiento:_gfVencimiento(b.periodo,'fin'), concepto:(b.concepto||'')+' (2ª quincena)'}));
  } else {
    run(b);
  }
  return {ok:creadas>0||dups>0, creadas:creadas, dup:dups>0&&creadas===0};
}

function programarGastoFijo(b) {
  try {
    var ss=SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egSh=ss.getSheetByName(EGRESOS_TABS[2026]||'Egresos2026');
    var iRec1=_egColEnsure(egSh,'recurrente','RecurrenteID');
    return _gfProgramarConFrec(egSh, iRec1, b, b.usuario||'sistema');
  } catch(ex){ return {ok:false, error:ex.message}; }
}

/* Regresa una CxP (fila de Egresos) que vino de un gasto fijo de vuelta a
   "Programación del mes": borra la fila para que el gasto fijo reaparezca como
   por programar. Solo si NO está pagada y sin abonos parciales. */
function regresarCxPaProgramacion(b) {
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egSh = ss.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
    var rowNum = parseInt(b.rowNum);
    if (!rowNum || rowNum < 2 || rowNum > egSh.getLastRow()) return { ok: false, error: 'Fila inválida' };
    var iRec1 = _egColEnsure(egSh, 'recurrente', 'RecurrenteID');
    var lastCol = egSh.getLastColumn();
    var row = egSh.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var recId = String(row[iRec1 - 1] || '').trim();
    if (!recId) return { ok: false, error: 'Esta cuenta no viene de un gasto fijo; no se puede regresar a programación.' };
    var cxpId = String(row[0] || '');
    var pagado = row[13] === true || String(row[13]).toUpperCase() === 'TRUE';
    if (pagado) return { ok: false, error: 'Esta cuenta ya está pagada. Cancélala primero antes de regresarla a programación.' };
    // Abonos parciales activos → no borrar sin reversar
    try { if (typeof _cxpSumAbonosActivos === 'function' && _cxpSumAbonosActivos(cxpId) > 0.01)
      return { ok: false, error: 'Esta cuenta tiene abonos parciales. Cancélala (⊘) para reversarlos antes de regresarla.' }; } catch (e) {}
    var proveedor = String(row[4] || ''), concepto = String(row[8] || ''), periodo = _gfToYM(row[2]);
    egSh.deleteRow(rowNum);
    try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (e) {}
    try { logAudit(b.usuario || 'sistema', 'GastoFijo', 'Regresar a programación', recId, 'Mes', periodo, proveedor + ' · ' + concepto); } catch (e) {}
    return { ok: true, recurrenteId: recId, periodo: periodo, proveedor: proveedor };
  } catch (ex) { return { ok: false, error: ex.message }; }
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
      var r=_gfProgramarConFrec(egSh, iRec1, items[i], b.usuario||'sistema');
      creadas+=r.creadas||0; if(r.dup)dups++; if(!r.ok&&!r.dup)errores++;
    }
    return {ok:true, creadas:creadas, duplicadas:dups, errores:errores};
  } catch(ex){ return {ok:false, error:ex.message}; }
}

// Diagnóstico manual (correr una vez desde el editor de Apps Script después del
// fix): busca en Egresos2026 combinaciones RecurrenteID+Mes repetidas — es decir,
// el mismo gasto fijo programado más de una vez para el mismo periodo, señal de
// pago duplicado. No borra nada, solo reporta para revisión manual.
function detectarDuplicadosGastosFijos() {
  var ss=SpreadsheetApp.openById(EGRESOS_SS_2026);
  var egSh=ss.getSheetByName(EGRESOS_TABS[2026]||'Egresos2026');
  var hdr=egSh.getRange(1,1,1,egSh.getLastColumn()).getValues()[0].map(function(h){return String(h).trim().toLowerCase();});
  var iRec=-1; for(var c=0;c<hdr.length;c++){ if(hdr[c].indexOf('recurrente')>-1){iRec=c;break;} }
  if(iRec<0) { Logger.log('No existe columna RecurrenteID todavía — no hay nada que revisar.'); return {ok:true, duplicados:[]}; }
  var data=egSh.getDataRange().getValues();
  var vistos={}, duplicados=[];
  for(var i=1;i<data.length;i++){
    var rec=String(data[i][iRec]||'').trim(); if(!rec) continue;
    var mes=_gfToYM(data[i][2]);
    var key=rec+'|'+mes;
    if(!vistos[key]) vistos[key]=[];
    vistos[key].push({fila:i+1, id:data[i][0], proveedor:data[i][4], concepto:data[i][8], monto:data[i][9], pagado:data[i][13]});
  }
  for(var k in vistos){
    if(vistos[k].length>1) duplicados.push({recurrenteYMes:k, filas:vistos[k]});
  }
  if(duplicados.length){
    Logger.log('⚠ Se encontraron '+duplicados.length+' gastos fijos duplicados en Egresos2026:');
    duplicados.forEach(function(d){
      Logger.log(d.recurrenteYMes+' → filas: '+d.filas.map(function(f){return 'fila '+f.fila+' (ID '+f.id+', '+f.proveedor+', $'+f.monto+', pagado='+f.pagado+')';}).join(' | '));
    });
  } else {
    Logger.log('✓ No se encontraron duplicados de RecurrenteID+Mes en Egresos2026.');
  }
  return {ok:true, duplicados:duplicados};
}

function bulkUpdateCxPMonto(body) {
  try {
    var rowNums=body.rowNums||[]; if(!rowNums.length) return {ok:false,error:'Sin filas'};
    var accion=String(body.accion||'set'); // 'set' | 'add'
    var valor=parseFloat(body.valor)||0;
    var anio=new Date().getFullYear();
    var ss=SpreadsheetApp.openById(EGRESOS_SS_2026);
    var sh=ss.getSheetByName(EGRESOS_TABS[anio]||'Egresos2026');
    if(!sh) return {ok:false,error:'Hoja no encontrada'};
    var COL_MONTO=10; // col J
    var updated=0;
    for(var i=0;i<rowNums.length;i++){
      var rn=parseInt(rowNums[i]); if(!rn||rn<2) continue;
      var cell=sh.getRange(rn,COL_MONTO);
      var cur=parseFloat(cell.getValue())||0;
      var nv=accion==='add'?cur+valor:valor;
      if(nv<0)nv=0;
      cell.setValue(Math.round(nv*100)/100);
      updated++;
    }
    try{CacheService.getScriptCache().remove('gas_egresos_v1_'+anio);}catch(e){}
    return {ok:true, updated:updated};
  } catch(ex){ return {ok:false,error:ex.message}; }
}
