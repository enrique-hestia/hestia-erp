/* ══════════════════════════════════════════════════════════════
   planes.gs — Planes de pago (calendario) para Cuentas por Cobrar
   --------------------------------------------------------------
   Un plan de pagos es un CALENDARIO INFORMATIVO para un adeudo de un paciente:
   parte el saldo en N pagos (quincenal o mensual) con fechas automáticas. NO
   mueve dinero ni toca bancos ni el saldo — el cobro se sigue haciendo con el
   botón "Cobrar / Abonar" de siempre. El calendario aparece en el Estado de
   Cuenta y marca cada pago como cubierto según lo abonado hasta la fecha.

   Hoja "Planes_Pago" en SHEET_ID. Un plan por (paciente, OP). Leer: cualquier
   sesión válida. Activar/desactivar: permiso editar_ingresos (gate en finance).
   ══════════════════════════════════════════════════════════════ */

var PLAN_TAB = 'Planes_Pago';
var PLAN_HEADERS = ['Id','Paciente','OP','MontoTotal','Frecuencia','NumPagos','FechaInicio','Activo','CreadoEn','ActualizadoEn'];

function setupPlanesPago() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(PLAN_TAB);
  if (!sh) sh = ss.insertSheet(PLAN_TAB);
  sh.clear();
  sh.getRange(1,1,1,PLAN_HEADERS.length).setValues([PLAN_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, PLAN_HEADERS.length, [110,240,110,110,100,80,110,70,150,150]);
  return {ok:true, tab:PLAN_TAB};
}
function _planSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(PLAN_TAB);
  if (!sh) { setupPlanesPago(); sh = ss.getSheetByName(PLAN_TAB); }
  return sh;
}
function _planNorm(s){ return String(s==null?'':s).trim(); }
function _planLow(s){ return _planNorm(s).toLowerCase(); }
function _planYmd(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
function _planParse(s){
  var m = String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(parseInt(m[1],10), parseInt(m[2],10)-1, parseInt(m[3],10)) : null;
}
function _planFmt(v){
  if (!v) return '';
  if (v instanceof Date) return _planYmd(v);
  return String(v).substring(0,10);
}
// Calendario: N pagos desde FechaInicio; quincenal=+15 días, mensual=+1 mes.
// El último pago absorbe el redondeo para que la suma == MontoTotal exacto.
function _planCalendario(montoTotal, frecuencia, numPagos, fechaInicio){
  var out = [];
  var base = _planParse(fechaInicio) || new Date();
  numPagos = Math.max(1, parseInt(numPagos,10) || 1);
  montoTotal = Math.round((parseFloat(montoTotal)||0)*100)/100;
  var cuota = Math.round((montoTotal/numPagos)*100)/100;
  var acum = 0;
  for (var k=0; k<numPagos; k++){
    var f;
    if (String(frecuencia).toLowerCase().indexOf('quinc') > -1){
      f = new Date(base.getTime()); f.setDate(f.getDate() + 15*k);
    } else { // mensual
      f = new Date(base.getFullYear(), base.getMonth()+k, base.getDate());
    }
    var monto = (k === numPagos-1) ? Math.round((montoTotal - acum)*100)/100 : cuota;
    acum = Math.round((acum + monto)*100)/100;
    out.push({ n:k+1, fecha:_planYmd(f), monto:monto });
  }
  return out;
}

// Lee los planes ACTIVOS de un paciente (o todos si no se da paciente), con su
// calendario ya calculado. El estado pagado/pendiente lo pinta el frontend con
// lo abonado (que ya conoce por el saldo).
function readPlanesPago(paciente){
  try {
    paciente = _planLow(paciente);
    var sh = _planSheet();
    var v = sh.getDataRange().getValues();
    var out = [];
    for (var i=1;i<v.length;i++){
      var r = v[i];
      var id = _planNorm(r[0]); if (!id) continue;
      var activo = (String(r[7]).toLowerCase() !== 'no' && String(r[7]).toLowerCase() !== 'false' && r[7] !== false);
      if (!activo) continue;
      if (paciente && _planLow(r[1]) !== paciente) continue;
      var montoTotal = parseFloat(r[3])||0, frec = _planNorm(r[4]), num = parseInt(r[5],10)||1, ini = _planFmt(r[6]);
      out.push({ id:id, paciente:_planNorm(r[1]), op:_planNorm(r[2]),
        montoTotal:montoTotal, frecuencia:frec, numPagos:num, fechaInicio:ini,
        calendario:_planCalendario(montoTotal, frec, num, ini) });
    }
    return { ok:true, planes:out };
  } catch(ex){ return { ok:false, error:ex.message, planes:[] }; }
}

// Activa (o reemplaza) el plan de un (paciente, OP). El dueño valida el permiso
// en el router. Upsert in-place; nunca sh.clear().
function activarPlanPagos(b){
  try {
    b = b || {};
    var paciente = _planNorm(b.paciente);
    if (!paciente) return { ok:false, error:'Falta el paciente.' };
    var op = _planNorm(b.op);
    var montoTotal = Number(b.montoTotal);
    if (!isFinite(montoTotal) || montoTotal <= 0) return { ok:false, error:'El monto del plan debe ser mayor a cero.' };
    montoTotal = Math.round(montoTotal*100)/100;
    var frec = _planLow(b.frecuencia);
    frec = (frec.indexOf('quinc') > -1) ? 'quincenal' : 'mensual';
    var numPagos = parseInt(b.numPagos,10);
    if (!(numPagos >= 1 && numPagos <= 120)) return { ok:false, error:'El plazo (número de pagos) debe estar entre 1 y 120.' };
    var ini = _planFmt(b.fechaInicio);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ini)) return { ok:false, error:'Fecha de inicio inválida (AAAA-MM-DD).' };

    var sh = _planSheet();
    var v = sh.getDataRange().getValues();
    var rowNum = -1, max = 0;
    for (var i=1;i<v.length;i++){
      var rid = _planNorm(v[i][0]); var m = rid.match(/(\d+)/); if (m){ var n=parseInt(m[1],10); if(n>max)max=n; }
      // Mismo (paciente, OP) activo → se reemplaza (un plan por adeudo).
      if (_planLow(v[i][1]) === paciente.toLowerCase() && _planNorm(v[i][2]) === op) rowNum = i+1;
    }
    var id = (rowNum>0) ? _planNorm(v[rowNum-1][0]) : ('PLAN-'+String(max+1).padStart(4,'0'));
    var creado = (rowNum>0) ? (v[rowNum-1][8] || new Date()) : new Date();
    var row = [ id, paciente, op, montoTotal, frec, numPagos, ini, true, creado, new Date() ];
    if (rowNum>0) sh.getRange(rowNum,1,1,PLAN_HEADERS.length).setValues([row]);
    else sh.appendRow(row);
    return { ok:true, id:id, calendario:_planCalendario(montoTotal, frec, numPagos, ini) };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

function desactivarPlanPago(b){
  try {
    b = b || {};
    var sh = _planSheet();
    var v = sh.getDataRange().getValues();
    var id = _planNorm(b.id), paciente = _planLow(b.paciente), op = _planNorm(b.op);
    for (var i=1;i<v.length;i++){
      var hit = id ? (_planNorm(v[i][0]) === id) : (_planLow(v[i][1]) === paciente && _planNorm(v[i][2]) === op);
      if (hit) { sh.getRange(i+1,8).setValue(false); sh.getRange(i+1,10).setValue(new Date()); return { ok:true }; }
    }
    return { ok:false, error:'No se encontró el plan.' };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
