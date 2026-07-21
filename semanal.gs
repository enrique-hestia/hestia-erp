/* ══════════════════════════════════════════════════════════════════════
   semanal.gs — Resumen / Dashboard Semanal para el Director
   Elige un día → semana [ref-6, ref] + comparación con la semana previa.
   Bloques A operación(metas) · B gasto por área · C KPIs · D ingresos/línea
   · E WoW · F caja/bancos · G CxP · H facturación.
   Reusa lectores existentes (summary/finance/board/presupuesto).
   Requiere: summary.gs, finance.gs, board.gs, presupuesto.gs en el proyecto.
   ══════════════════════════════════════════════════════════════════════ */

var SEMANAL_CFG_TAB = 'Semanal_Config';

function _semDate(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
function _semAddDays(iso, n){ var d=new Date(String(iso).substring(0,10)+'T12:00:00'); d.setDate(d.getDate()+n); return _semDate(d); }
function _semInRange(f, a, b){ return f && f>=a && f<=b; }
function _semNum(v){ if(typeof v==='number') return v; var n=parseFloat(String(v||'').replace(/[$,\s]/g,'')); return isNaN(n)?0:n; }

// Tipo de procedimiento para la operación semanal. SÍ incluye Ciclos Externos
// (a diferencia del Board Deck). Excluye estudios/medicamentos/insumos.
function _semTipoOperacion(nombre){
  var n=String(nombre||'').toLowerCase();
  if(/estudio|medicamento|laborator|almacen|suplemento|insumo/.test(n)) return null;
  if(/surrogacy/.test(n)) return 'Surrogacy';
  if(/reprovid/.test(n)) return 'Reprovida';
  if(/extern/.test(n))   return 'Ciclos Externos';
  if(/consulta/.test(n)) return 'Consulta';
  if(/baja/.test(n))     return 'Ciclo Baja';
  if(/alta/.test(n))     return 'Ciclo Alta';
  return null;
}

// Meta semanal por tipo = (meta trimestral ÷ 13 semanas) con override manual del config.
function _semMetaSemanal(ref, cfg){
  var out={};
  try{
    var y=parseInt(ref.substring(0,4),10), mo=parseInt(ref.substring(5,7),10), q=Math.ceil(mo/3);
    var pr=(typeof readPresupuesto==='function') ? readPresupuesto(y+'-Q'+q) : null;
    if(pr && pr.ok && pr.siguiente && pr.siguiente.ingresosGrupos){
      pr.siguiente.ingresosGrupos.forEach(function(g){
        var t=_semTipoOperacion(g.grupo); if(!t) return;
        out[t]=(out[t]||0) + (Number(g.cantProy)||0)/13;
      });
    }
  }catch(e){}
  // redondear autos a 1 decimal
  Object.keys(out).forEach(function(k){ out[k]=Math.round(out[k]*10)/10; });
  // overrides manuales
  var ov=(cfg && cfg.metasOverride) || {};
  Object.keys(ov).forEach(function(k){ var v=parseFloat(ov[k]); if(!isNaN(v)&&v>=0) out[k]=v; });
  return out;
}

function readResumenSemanal(fechaRef){
  try{
    var ref = String(fechaRef||'').substring(0,10) || _semDate(new Date());
    var ini = _semAddDays(ref, -6), fin = ref;
    var pIni = _semAddDays(ref, -13), pFin = _semAddDays(ref, -7);
    var cfg = readSemanalConfig();

    // Años a leer (por si la semana o su previa cruzan el borde de año)
    var years = {}; [ini, fin, pIni, pFin].forEach(function(x){ years[parseInt(x.substring(0,4),10)]=1; });
    years = Object.keys(years).map(Number);

    // ── Ingresos (operación + revenue por línea) ──
    var A={}, Aprev={}, revLinea={}, totIng=0, totIngPrev=0, nProc=0, nProcPrev=0;
    years.forEach(function(anio){
      var ing = (typeof _summaryReadIngresos==='function') ? _summaryReadIngresos(anio) : [];
      ing.forEach(function(r){
        var f=r.fecha; if(!f) return;
        var linea=r.grupoU||r.categoria||'—';
        var tipo=_semTipoOperacion(r.grupoU||r.categoria);
        if(_semInRange(f, ini, fin)){
          totIng+=r.total; revLinea[linea]=(revLinea[linea]||0)+r.total;
          if(tipo){ A[tipo]=(A[tipo]||0)+r.cantidad; nProc+=r.cantidad; }
        } else if(_semInRange(f, pIni, pFin)){
          totIngPrev+=r.total;
          if(tipo){ Aprev[tipo]=(Aprev[tipo]||0)+r.cantidad; nProcPrev+=r.cantidad; }
        }
      });
    });

    // ── Egresos por área (subtipo) ──
    var gastoArea={}, totEg=0, totEgPrev=0;
    years.forEach(function(anio){
      var eg = (typeof readEgresosData==='function') ? readEgresosData(anio, {rowsOnly:true, skipUrls:true}) : {rows:[]};
      (eg.rows||[]).forEach(function(r){
        var f=(r.fecha||'').substring(0,10); if(!f) return;
        if(r.estatus==='Cancelada') return;
        var m=r.monto||0, area=(String(r.subtipo||'').trim())||'Otros';
        if(_semInRange(f, ini, fin)){ totEg+=m; gastoArea[area]=(gastoArea[area]||0)+m; }
        else if(_semInRange(f, pIni, pFin)){ totEgPrev+=m; }
      });
    });

    // ── A. Operación: meta vs real (avance/faltante) ──
    var metaSemanal = _semMetaSemanal(ref, cfg);
    var tipos={}; Object.keys(metaSemanal).forEach(function(k){tipos[k]=1;}); Object.keys(A).forEach(function(k){tipos[k]=1;});
    var operacion = Object.keys(tipos).map(function(t){
      var real=A[t]||0, meta=metaSemanal[t]||0;
      var avance = meta>0 ? (real/meta*100) : (real>0?100:0);
      return { tipo:t, real:real, meta:meta, avancePct:avance, faltantePct:Math.max(0,100-avance), realPrev:Aprev[t]||0 };
    }).sort(function(a,b){ return (b.meta-a.meta) || (b.real-a.real); });

    // ── B. Gasto por área ──
    var gastoAreas = Object.keys(gastoArea).map(function(k){ return {area:k, monto:gastoArea[k]}; }).sort(function(a,b){ return b.monto-a.monto; });

    // ── D. Ingresos por línea ──
    var ingresosLineas = Object.keys(revLinea).map(function(k){ return {linea:k, monto:revLinea[k]}; }).sort(function(a,b){ return b.monto-a.monto; });

    // ── C / E. KPIs + WoW ──
    var util=totIng-totEg, utilPrev=totIngPrev-totEgPrev;
    function wow(a,b){ return b ? ((a-b)/Math.abs(b)*100) : null; }
    var kpis={
      ingresos:totIng, egresos:totEg, utilidad:util, margen: totIng? util/totIng*100 : 0,
      procedimientos:nProc, ticket: nProc? totIng/nProc : 0,
      ingresosPrev:totIngPrev, egresosPrev:totEgPrev, utilidadPrev:utilPrev, procedimientosPrev:nProcPrev,
      wowIngresos:wow(totIng,totIngPrev), wowEgresos:wow(totEg,totEgPrev), wowUtilidad:wow(util,utilPrev), wowProc:wow(nProc,nProcPrev)
    };

    // ── F. Caja / bancos (movimientos ya filtrados por rango) ──
    var caja=[];
    try{
      var bk=readBanksData(ini, fin);
      var bancos=bk.bancos||{};
      Object.keys(bancos).forEach(function(key){
        var b=bancos[key]; if(!b) return;
        var ent=0, sal=0;
        (b.movimientos||[]).forEach(function(m){
          var f=(m.fecha||'').substring(0,10); if(!_semInRange(f, ini, fin)) return;
          var flow = (_semNum(m.deposito)-_semNum(m.retiro)) || _semNum(m.monto) || _semNum(m.totalVenta);
          if(flow>0) ent+=flow; else sal+=Math.abs(flow);
        });
        caja.push({ banco:b.nombre||key, entradas:ent, salidas:sal, saldo:b.saldo||0 });
      });
    }catch(e){}

    // ── G. CxP (vence esta semana / próxima / total pendiente) ──
    var cxp={ venceSemana:0, totalSemana:0, venceProx:0, totalProx:0, totalPendiente:0 };
    try{
      var bd=readBDCxP();
      var proxIni=_semAddDays(fin,1), proxFin=_semAddDays(fin,7);
      (bd.rows||[]).forEach(function(r){
        var v=(r.vencimiento||'').substring(0,10), m=r.monto||0;
        cxp.totalPendiente+=m;
        if(_semInRange(v, ini, fin)){ cxp.venceSemana++; cxp.totalSemana+=m; }
        else if(_semInRange(v, proxIni, proxFin)){ cxp.venceProx++; cxp.totalProx+=m; }
      });
    }catch(e){}

    // ── H. Facturación de la semana (best-effort desde comprobantes) ──
    var facturacion=null;
    try{
      if(typeof readComprobantesRango==='function'){
        var fc=readComprobantesRango(ini, fin);
        if(fc && fc.ok) facturacion={ count:fc.count||0, total:fc.total||0 };
      }
    }catch(e){}

    return { ok:true,
      periodo:{inicio:ini, fin:fin, ref:ref},
      periodoPrev:{inicio:pIni, fin:pFin},
      config:cfg, kpis:kpis, operacion:operacion, gastoAreas:gastoAreas,
      ingresosLineas:ingresosLineas, caja:caja, cxp:cxp, facturacion:facturacion,
      metaSemanal:metaSemanal };
  }catch(ex){ return {ok:false, error:ex.message+' (L'+(ex.lineNumber||'?')+')'}; }
}

/* ── Config: bloques on/off + meta semanal override ── */
function readSemanalConfig(){
  var def={ bloques:{A:true,B:true,C:true,D:true,E:true,F:true,G:true,H:true}, metasOverride:{} };
  try{
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var sh=ss.getSheetByName(SEMANAL_CFG_TAB);
    if(!sh) return def;
    var data=sh.getDataRange().getValues();
    var cfg={ bloques:{A:true,B:true,C:true,D:true,E:true,F:true,G:true,H:true}, metasOverride:{} };
    for(var i=1;i<data.length;i++){
      var k=String(data[i][0]||'').trim(), v=data[i][1];
      if(!k) continue;
      if(k.indexOf('bloque:')===0){ cfg.bloques[k.substring(7)] = (v===true||String(v).toLowerCase()==='true'||v==='1'||v===1); }
      else if(k.indexOf('meta:')===0){ var n=parseFloat(v); if(!isNaN(n)) cfg.metasOverride[k.substring(5)]=n; }
    }
    return cfg;
  }catch(e){ return def; }
}
function saveSemanalConfig(body){
  try{
    var ss=SpreadsheetApp.openById(SHEET_ID);
    var sh=ss.getSheetByName(SEMANAL_CFG_TAB) || ss.insertSheet(SEMANAL_CFG_TAB);
    sh.clear();
    sh.getRange(1,1,1,2).setValues([['Clave','Valor']]);
    var rows=[];
    var bl=body.bloques||{};
    Object.keys(bl).forEach(function(k){ rows.push(['bloque:'+k, !!bl[k]]); });
    var mo=body.metasOverride||{};
    Object.keys(mo).forEach(function(k){ var n=parseFloat(mo[k]); if(!isNaN(n)&&mo[k]!=='') rows.push(['meta:'+k, n]); });
    if(rows.length) sh.getRange(2,1,rows.length,2).setValues(rows);
    try{ logAudit(body.usuario||'sistema','Semanal','Guardar config','','','',''); }catch(e){}
    return {ok:true};
  }catch(ex){ return {ok:false, error:ex.message}; }
}
function setupSemanalConfig(){ return saveSemanalConfig({bloques:{A:true,B:true,C:true,D:true,E:true,F:true,G:true,H:true}, metasOverride:{}}); }
