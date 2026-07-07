/* ══════════════════════════════════════════════════════════════
   board.gs — Board Deck (Informe Ejecutivo de Resultados)
   ------------------------------------------------------------
   Arma el paquete de datos para la presentación ejecutiva a socios:
   P&L (readSummary) + serie mensual del año + meta/presupuesto +
   KPIs financieros + narrativa automática según el periodo elegido.
   No depende de librerías externas; el frontend grafica con Chart.js.
   Requiere: summary.gs, finance.gs (readEgresosData, _summaryReadIngresos).
   ══════════════════════════════════════════════════════════════ */

var BOARD_CFG_TAB = 'Board_Config';

function _boardNum(v){ return Number(v)||0; }
function _boardPct(a,b){ return b>0 ? a/b : 0; }

/* Tipo de periodo según el rango (mes / trimestre / semestre / año / rango). */
function _boardPeriodType(fi, ff){
  var di=new Date(fi+'T12:00:00'), df=new Date(ff+'T12:00:00');
  var dias = Math.round((df-di)/86400000)+1;
  var mismoAnio = fi.substring(0,4)===ff.substring(0,4);
  var iniMes = fi.substring(8,10)==='01';
  if (dias>=28 && dias<=31 && fi.substring(0,7)===ff.substring(0,7)) return {tipo:'mes', label:_boardMesLabel(fi)};
  if (dias>=88 && dias<=93) return {tipo:'trimestre', label:_boardTrimestre(fi)};
  if (dias>=180 && dias<=185) return {tipo:'semestre', label:(di.getMonth()<6?'1er':'2do')+' semestre '+fi.substring(0,4)};
  if (dias>=360 && dias<=366) return {tipo:'año', label:'Año '+fi.substring(0,4)};
  return {tipo:'rango', label:fi+' a '+ff};
}
var _BOARD_MESES=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
function _boardMesLabel(iso){ var m=parseInt(iso.substring(5,7),10)-1; return (_BOARD_MESES[m]||'')+' '+iso.substring(0,4); }
function _boardTrimestre(iso){ var m=parseInt(iso.substring(5,7),10); var q=Math.floor((m-1)/3)+1; return 'Q'+q+' '+iso.substring(0,4); }

/* Serie mensual del año: {mes, mesIdx, ingresos, egresos, neto}. */
function _boardSeries(year){
  var meses = [];
  for (var i=0;i<12;i++) meses.push({mes:_BOARD_MESES[i].substring(0,3), mesIdx:i, ingresos:0, egresos:0, neto:0});
  // Ingresos
  try {
    var ins = _summaryReadIngresos(year);
    ins.forEach(function(r){
      var f=(r.fecha||'').substring(0,10);
      if (f.substring(0,4)!==String(year)) return;
      var mi=parseInt(f.substring(5,7),10)-1;
      if (mi>=0 && mi<12) meses[mi].ingresos += _boardNum(r.total);
    });
  } catch(e){}
  // Egresos (excluye Cancelada y Crédito/AMEX del P&L)
  try {
    var eg = readEgresosData(year);
    (eg.rows||[]).forEach(function(r){
      if (r.estatus==='Cancelada') return;
      if (_sumNorm(r.contable)==='credito' || _sumNorm(r.contable)==='crédito') return;
      var f=(r.fecha||r.vencimiento||'').substring(0,10);
      if (f.substring(0,4)!==String(year)) return;
      var mi=parseInt(f.substring(5,7),10)-1;
      if (mi>=0 && mi<12) meses[mi].egresos += _boardNum(r.monto);
    });
  } catch(e){}
  meses.forEach(function(m){ m.neto = m.ingresos - m.egresos; });
  return meses;
}

/* Meta / presupuesto del periodo (best-effort; null si no hay presupuesto.gs). */
function _boardMeta(fi, ff){
  try {
    if (typeof readPresupuesto !== 'function') return null;
    var p = readPresupuesto(fi, ff);
    if (!p || !p.ok) return null;
    // Estructura tolerante: busca metas de ingresos/egresos/utilidad
    return {
      ingresosMeta: _boardNum(p.metaIngresos || p.ingresosMeta || (p.meta && p.meta.ingresos) || 0),
      egresosMeta:  _boardNum(p.metaEgresos  || p.egresosMeta  || (p.meta && p.meta.egresos)  || 0),
      utilidadMeta: _boardNum(p.metaUtilidad || p.utilidadMeta || (p.meta && p.meta.utilidad) || 0)
    };
  } catch(e){ return null; }
}

/* Narrativa automática, en lenguaje simple, adaptada al tipo de periodo. */
function _boardNarrative(sum, kpis, per, serie){
  var fmt = function(v){ return '$'+Math.round(v).toLocaleString('en-US'); };
  var pctTxt = function(v){ return (v>=0?'+':'')+(v*100).toFixed(1)+'%'; };
  var partes = [];
  var pl = per.tipo==='mes'?'el mes de '+per.label : (per.tipo==='trimestre'?'el trimestre '+per.label : (per.tipo==='año'?'el año '+sum.periodo.inicio.substring(0,4) : 'el periodo'));
  // Ingresos
  var gRev = kpis.crecimientoIngresos;
  partes.push('Durante '+pl+', la clínica generó ingresos por '+fmt(kpis.ingresos)+
    (isFinite(gRev)&&kpis.ingresosPrev>0 ? ' ('+pctTxt(gRev)+' vs. el periodo anterior)':'')+'.');
  // Top línea de ingreso
  var revLines=(sum.lineas||[]).filter(function(x){return x.tipo==='dato'&&x.grupo==='REVENUE';}).sort(function(a,b){return b.actual-a.actual;});
  if (revLines.length){
    var top=revLines[0], low=revLines[revLines.length-1];
    partes.push('La línea que más aportó fue '+top.linea+' con '+fmt(top.actual)+' ('+(kpis.ingresos>0?(top.actual/kpis.ingresos*100).toFixed(1):'0')+'% del total)'+
      (revLines.length>1?', mientras que '+low.linea+' fue la de menor contribución':'')+'.');
  }
  // Egresos y utilidad
  partes.push('Los egresos operativos sumaron '+fmt(kpis.egresos)+', dejando una utilidad neta de '+fmt(kpis.utilidadNeta)+
    ' (margen neto de '+(kpis.margenNeto*100).toFixed(1)+'%).');
  // Meta
  if (kpis.cumplimientoMeta!=null){
    var cm=kpis.cumplimientoMeta;
    partes.push(cm>=1 ? 'Se superó la meta de ingresos en '+pctTxt(cm-1)+' ('+fmt(kpis.ingresos-kpis.ingresosMeta)+' por encima).'
      : 'Se alcanzó el '+(cm*100).toFixed(0)+'% de la meta de ingresos (faltaron '+fmt(kpis.ingresosMeta-kpis.ingresos)+').');
  }
  // Tendencia del año
  if (serie && serie.length){
    var conDatos=serie.filter(function(m){return m.ingresos>0||m.egresos>0;});
    if (conDatos.length>=2){
      var mejor=conDatos.slice().sort(function(a,b){return b.neto-a.neto;})[0];
      partes.push('En lo que va del año, '+_BOARD_MESES[mejor.mesIdx]+' fue el mes de mayor utilidad ('+fmt(mejor.neto)+').');
    }
  }
  return partes.join(' ');
}

/* ── El paquete completo para el Board Deck ─────────────────────── */
function readBoardReport(fechaInicio, fechaFin){
  try {
    var sum = readSummary(fechaInicio, fechaFin);
    if (!sum.ok) return sum;
    var fi=sum.periodo.inicio, ff=sum.periodo.fin;
    var year=parseInt(ff.substring(0,4),10);
    var m=sum.metricas, mp=sum.metricasPrev||{};
    var rev=_boardNum(m.revenue);
    var egr=_boardNum(sum.reconc && sum.reconc.egresosPL);
    if(!egr) egr=_boardNum(m.cogs)+_boardNum(m.opex)+_boardNum(m.ga)+_boardNum(m.taxes);
    var net=_boardNum(m.netProfit);
    var revPrev=_boardNum(mp.revenue), netPrev=_boardNum(mp.netProfit);
    var egrPrev=_boardNum(mp.cogs)+_boardNum(mp.opex)+_boardNum(mp.ga)+_boardNum(mp.taxes);

    var meta=_boardMeta(fi,ff);
    var serie=_boardSeries(year);
    var per=_boardPeriodType(fi,ff);

    // # ciclos y ticket (de sub-items de REVENUE con cantidad)
    var ciclos=0;
    (sum.lineas||[]).forEach(function(l){ if(l.tipo==='dato'&&l.grupo==='REVENUE'){ (l.subitems||[]).forEach(function(s){
      if(s.productos){ s.productos.forEach(function(p){ ciclos+=_boardNum(p.cantidad); }); } else { ciclos+=_boardNum(s.cantidad); }
    }); } });

    var kpis={
      ingresos:rev, egresos:egr, utilidadNeta:net,
      margenNeto:_boardPct(net,rev), margenBruto:_boardPct(_boardNum(m.grossProfit),rev),
      ebitda:_boardNum(m.ebitda), margenEbitda:_boardPct(_boardNum(m.ebitda),rev),
      grossProfit:_boardNum(m.grossProfit),
      ingresosPrev:revPrev, egresosPrev:egrPrev, utilidadPrev:netPrev,
      crecimientoIngresos:_boardPct(rev-revPrev,revPrev),
      crecimientoEgresos:_boardPct(egr-egrPrev,egrPrev),
      crecimientoUtilidad:_boardPct(net-netPrev,netPrev),
      ciclos:ciclos, ticketPromedio: ciclos>0?rev/ciclos:0,
      ingresosMeta: meta?meta.ingresosMeta:null,
      cumplimientoMeta: (meta&&meta.ingresosMeta>0)?_boardPct(rev,meta.ingresosMeta):null
    };

    var narrativa=_boardNarrative(sum, kpis, per, serie);

    // Composición para donas
    var compIngresos=(sum.lineas||[]).filter(function(x){return x.tipo==='dato'&&x.grupo==='REVENUE';})
      .map(function(l){return {label:l.linea, valor:l.actual};}).sort(function(a,b){return b.valor-a.valor;});
    var compEgresos=[
      {label:'COGS', valor:_boardNum(m.cogs)},
      {label:'OpEx', valor:_boardNum(m.opex)},
      {label:'G&A', valor:_boardNum(m.ga)},
      {label:'Taxes', valor:_boardNum(m.taxes)}
    ].filter(function(x){return x.valor>0;});

    return {
      ok:true, periodo:sum.periodo, prev:sum.prev, tipoPeriodo:per,
      kpis:kpis, narrativa:narrativa, serie:serie, meta:meta,
      compIngresos:compIngresos, compEgresos:compEgresos,
      lineas:sum.lineas, reconc:sum.reconc, amexCredito:sum.amexCredito, metricas:sum.metricas
    };
  } catch(ex){ return { ok:false, error:ex.message+' (L:'+ex.lineNumber+')' }; }
}

/* ── Config de la plantilla (portada, colores, título) ──────────── */
function readBoardConfig(){
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var sh=ss.getSheetByName(BOARD_CFG_TAB);
    if(!sh){ return { ok:true, config:_boardDefaultConfig() }; }
    var data=sh.getDataRange().getValues();
    var cfg=_boardDefaultConfig();
    for(var i=1;i<data.length;i++){ var k=String(data[i][0]||'').trim(); if(k) cfg[k]=String(data[i][1]||''); }
    return { ok:true, config:cfg };
  } catch(ex){ return { ok:false, error:ex.message, config:_boardDefaultConfig() }; }
}
function _boardDefaultConfig(){
  return {
    titulo:'Resultados', empresa:'Hestia Fertility', subtitulo:'Informe Ejecutivo de Resultados',
    coverUrl:'', logoUrl:'',
    colorPrimario:'#6b6b6b', colorAcento:'#c46a7a',
    colorIngresos:'#8a9a8f', colorEgresos:'#c8969c', colorFondo:'#f6f4f1'
  };
}
function saveBoardConfig(body){
  try {
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var sh=ss.getSheetByName(BOARD_CFG_TAB);
    if(!sh){ sh=ss.insertSheet(BOARD_CFG_TAB); sh.getRange(1,1,1,2).setValues([['Clave','Valor']]).setFontWeight('bold'); }
    var cfg=body.config||{};
    var last=sh.getLastRow(); if(last>1) sh.getRange(2,1,last-1,2).clearContent();
    var rows=Object.keys(cfg).map(function(k){ return [k, String(cfg[k]==null?'':cfg[k])]; });
    if(rows.length) sh.getRange(2,1,rows.length,2).setValues(rows);
    return { ok:true, guardadas:rows.length };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
