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

/* Serie de los ÚLTIMOS 12 MESES terminando en el mes del periodo (un año hacia
   atrás), con el valor del mismo mes del año anterior para comparar (YoY).
   Devuelve [{label:'Jun 25', anio, mesIdx, ingresos, egresos, neto,
              ingresosPrevAnio, egresosPrevAnio}]. */
function _boardSeries(endIso){
  var endY=parseInt(endIso.substring(0,4),10), endM=parseInt(endIso.substring(5,7),10)-1; // 0-based
  var MES3=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var buckets={}; // 'YYYY-MM' -> {ing, egr}
  function add(key, field, val){ if(!buckets[key]) buckets[key]={ing:0,egr:0}; buckets[key][field]+=val; }
  // Leer 3 años para cubrir la ventana móvil + el overlay del año anterior
  for (var y=endY-2; y<=endY; y++){
    try { _summaryReadIngresos(y).forEach(function(r){ var f=(r.fecha||'').substring(0,7); if(f.length===7) add(f,'ing',_boardNum(r.total)); }); } catch(e){}
    try {
      var eg=readEgresosData(y);
      (eg.rows||[]).forEach(function(r){
        if(r.estatus==='Cancelada') return;
        if(_sumNorm(r.contable)==='credito'||_sumNorm(r.contable)==='crédito') return;
        var f=(r.fecha||r.vencimiento||'').substring(0,7); if(f.length===7) add(f,'egr',_boardNum(r.monto));
      });
    } catch(e2){}
  }
  var out=[];
  for (var i=11;i>=0;i--){
    var d=new Date(endY, endM-i, 1);
    var key=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
    var pk=(d.getFullYear()-1)+'-'+String(d.getMonth()+1).padStart(2,'0');
    var b=buckets[key]||{ing:0,egr:0}, p=buckets[pk]||{ing:0,egr:0};
    out.push({ label:MES3[d.getMonth()]+" '"+String(d.getFullYear()).slice(-2), anio:d.getFullYear(), mesIdx:d.getMonth(),
      ingresos:b.ing, egresos:b.egr, neto:b.ing-b.egr, ingresosPrevAnio:p.ing, egresosPrevAnio:p.egr });
  }
  return out;
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
  // Ingresos (solo se menciona crecimiento si hay periodo anterior con datos)
  var gRev = kpis.crecimientoIngresos;
  partes.push('Durante '+pl+', la clínica generó ingresos por '+fmt(kpis.ingresos)+
    (isFinite(gRev)&&kpis.ingresosPrev>0 ? ' ('+pctTxt(gRev)+' vs. el periodo anterior)':'')+'.');
  // Top/bottom línea de ingreso — SOLO líneas con datos (> 0)
  var revLines=(sum.lineas||[]).filter(function(x){return x.tipo==='dato'&&x.grupo==='REVENUE'&&x.actual>0;}).sort(function(a,b){return b.actual-a.actual;});
  if (revLines.length){
    var top=revLines[0];
    var frase='La línea que más aportó fue '+top.linea+' con '+fmt(top.actual)+' ('+(kpis.ingresos>0?(top.actual/kpis.ingresos*100).toFixed(1):'0')+'% del total)';
    if (revLines.length>=3){ var low=revLines[revLines.length-1]; frase+=', mientras que '+low.linea+' fue la de menor aporte con '+fmt(low.actual); }
    partes.push(frase+'.');
  }
  // Egresos y resultado
  var perdida = kpis.utilidadNeta < 0;
  if (perdida){
    partes.push('Los egresos operativos sumaron '+fmt(kpis.egresos)+', mayores a los ingresos, dejando una pérdida de '+fmt(Math.abs(kpis.utilidadNeta))+' (margen de '+(kpis.margenNeto*100).toFixed(1)+'%).');
    // ¿Por qué alcanzó el dinero para pagar pese a la pérdida?
    var razones=[];
    if (kpis.utilidadPrev > 0) razones.push('el remanente del periodo anterior, que cerró con una utilidad de '+fmt(kpis.utilidadPrev));
    var amex = (sum.amexCredito && sum.amexCredito.actual) || 0;
    if (amex > 0) razones.push('el apalancamiento con la tarjeta AMEX ('+fmt(amex)+'), cuyo pago se difiere al siguiente periodo');
    if (razones.length) partes.push('Aun con el resultado negativo, la operación pudo cubrir sus pagos gracias a '+razones.join(' y ')+'.');
    else partes.push('La diferencia se cubrió con la caja disponible de la clínica.');
  } else {
    partes.push('Los egresos operativos sumaron '+fmt(kpis.egresos)+', dejando una utilidad neta de '+fmt(kpis.utilidadNeta)+
      ' (margen neto de '+(kpis.margenNeto*100).toFixed(1)+'%).');
  }
  // Meta (solo si existe)
  if (kpis.cumplimientoMeta!=null && kpis.ingresosMeta>0){
    var cm=kpis.cumplimientoMeta;
    partes.push(cm>=1 ? 'Se superó la meta de ingresos en '+pctTxt(cm-1)+' ('+fmt(kpis.ingresos-kpis.ingresosMeta)+' por encima).'
      : 'Se alcanzó el '+(cm*100).toFixed(0)+'% de la meta de ingresos (faltaron '+fmt(kpis.ingresosMeta-kpis.ingresos)+').');
  }
  // Mejor mes (solo con datos suficientes)
  if (serie && serie.length){
    var conDatos=serie.filter(function(m){return m.ingresos>0||m.egresos>0;});
    if (conDatos.length>=2){
      var mejor=conDatos.slice().sort(function(a,b){return b.neto-a.neto;})[0];
      partes.push('En los últimos doce meses, '+_BOARD_MESES[mejor.mesIdx]+' fue el mes de mayor utilidad ('+fmt(mejor.neto)+').');
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
    var serie=_boardSeries(ff);
    var per=_boardPeriodType(fi,ff);

    // Detalle de egresos: secciones (COGS/OpEx/G&A/Taxes) con sus subgrupos y proveedores
    var egSecs={COGS:'COGS',OPEX:'OpEx',GA:'G&A',TAXES:'Taxes'};
    var egresosDetalle=[];
    (sum.lineas||[]).forEach(function(l){
      if(l.tipo==='seccion' && egSecs[l.grupo]) egresosDetalle.push({nivel:'seccion',label:l.label,valor:l.actual});
      else if(l.tipo==='dato' && egSecs[l.grupo]){
        egresosDetalle.push({nivel:'sub',label:l.linea,valor:l.actual,
          proveedores:(l.subitems||[]).slice(0,6).map(function(s){return {label:s.label,valor:s.actual};})});
      }
    });

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
      compIngresos:compIngresos, compEgresos:compEgresos, egresosDetalle:egresosDetalle,
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
