/* ══════════════════════════════════════════════════════════════
   api_config.gs — Constantes globales del ERP Hestia
   ──────────────────────────────────────────────────────────────
   INSTRUCCIONES: Copiar este archivo como un .gs separado dentro
   del mismo proyecto de Google Apps Script. Todas las constantes
   aquí son accesibles globalmente por api_core.gs y api_finance.gs
   ══════════════════════════════════════════════════════════════ */

var API_VERSION  = 'v2026-06-15';
var AUTH_SECRET  = 'hestia2026erp-secret'; // ← Cambia este valor por algo único y secreto

/* ── IDs de Spreadsheets ──────────────────────────────────── */
var SHEET_ID    = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4'; // ERP principal
var ER_SS_ID    = '17jlXzaIvohpN_UoE2kvLK1Bb6P6JxyKONsnrDX_St9U'; // Estado de Resultados
var BANKS_SS_ID = '1O1tmtuVMlDl6rsN0IVFH14KmYjQZhY6GOs68PU_u0cg'; // Cuentas bancarias
var LAB_SS_ID   = '1hYmIl4gSTVrvghP7KY0y0dC200o8w0zShXj63zP-TrQ'; // Laboratorio
var MED_SS_ID   = '1fiuUtw-sg2ELNxq9bCjaOtRz1n87wuVi8IOQYzEi8tM'; // Medicamentos
var PAC_SS_ID   = '1uoQU-vbefxWwaLxJyTFT25gj7Nr2223WISa3tqH-Rio'; // Pacientes
var PROD_SS_ID  = '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY'; // Productos
var QX_SS_ID    = LAB_SS_ID; // ← Reemplaza cuando exista el sheet de Quirófano
var CAJA_CHICA_SS_ID = '1uB9HnQLqHbotP0w21z6mVQcc4ABb428iKXEru8hvDQE'; // Caja Chica

/* ── GIDs (pestañas individuales por ID numérico) ─────────── */
var ER_GID     = 1953492149; // Estado de Resultados
var BUDGET_GID = 2097864117; // Budget
var PL_GID     = 1415550816; // Operating P&L Statement (referencia visual)

var BANKS_GID = {
  santander:   0,
  amex:        13958125,
  mercadopago: 1036684249
};

/* ── Mapeo: nombre de pestaña → spreadsheet donde vive ─────── */
var CAPTURA_SHEETS = {
  // Medicamentos
  'Medicamentos':    MED_SS_ID,
  'Orden_Compra':    MED_SS_ID,
  'Ent. Med':        MED_SS_ID,
  'Lista Med':       MED_SS_ID,
  'Estimulacion':    MED_SS_ID,
  'Estimulación':    MED_SS_ID,
  'Salidas Med':     MED_SS_ID,
  // Laboratorio
  'Resumen':         LAB_SS_ID,
  'ART Lab':         LAB_SS_ID,
  'FET':             LAB_SS_ID,
  'Andrología':      LAB_SS_ID,
  'Andrologia':      LAB_SS_ID,
  'Inventario Crío': LAB_SS_ID,
  'Inventario Crio': LAB_SS_ID,
  'Insumos':         LAB_SS_ID,
  'Salidas Lab':     LAB_SS_ID,
  // Quirófano
  'Insumos Qx':      QX_SS_ID,
  'Salidas Qx':      QX_SS_ID,
  // Otras
  'Pacientes':       PAC_SS_ID,
  'Productos':       PROD_SS_ID,
  'Caja Chica':      CAJA_CHICA_SS_ID
  // ── FASE 2: agregar aquí CxC, CxP, Nómina, etc. ───────────
};

/* ── Aliases: nombre alternativo → nombre exacto en Sheets ─── */
var SHEET_ALIASES = {
  'Orden_Compra': 'Ent. Med',
  'Estimulacion': 'Estimulación'
};

var CAPTURA_SHEET_ID_DEFAULT = SHEET_ID;

/* ══════════════════════════════════════════════════════════════
   api_finance.gs — Módulo Financiero del ERP Hestia
   ──────────────────────────────────────────────────────────────
   Funciones:
     · readCashFlow       — Flujo mensual desde hoja CashFlow
     · readServicios      — Servicios y márgenes
     · readFunnel         — Embudo de conversión
     · readAlertas        — Alertas del dashboard
     · readDonut          — Distribución donut
     · readPaisesOrigen   — Turismo médico
     · readCostos         — Distribución de costos
     · readEstadoResultados — P&L detallado desde hoja ER
     · readOperatingPL    — Operating P&L Statement (stateless)
     · _buildPLReport     — Constructor interno del P&L
     · readBanksData      — Saldos y movimientos de 3 bancos
     · saveBankRow        — Escribe movimiento en pestaña bancaria
     · doPost             — Handler HTTP POST
   ══════════════════════════════════════════════════════════════ */

/* Columnas CashFlow: A=Sucursal | B=Fecha | C=Mes | D=Flujo_MXN */
function readCashFlow(ss, fechaInicio, fechaFin, sucursal) {
  var hoja = ss.getSheetByName('CashFlow');
  if (!hoja) return { meses:[], flujo:[] };
  var rows = hoja.getDataRange().getValues();
  if (rows.length < 2) return { meses:[], flujo:[] };
  var hdrs = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iSuc   = hdrs.indexOf('sucursal');
  var iFecha = hdrs.indexOf('fecha');   if (iFecha < 0) iFecha = 1;
  var iMes   = hdrs.indexOf('mes');     if (iMes   < 0) iMes   = 2;
  var iFlujo = hdrs.indexOf('flujo_mxn');
  if (iFlujo < 0) iFlujo = hdrs.indexOf('flujo');
  if (iFlujo < 0) iFlujo = 3;
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[iFecha]).trim();
    if (f < fechaInicio || f > fechaFin) return false;
    if (sucursal && sucursal !== 'Todas' && iSuc >= 0) {
      var s = String(r[iSuc] || '').trim();
      if (s && s !== sucursal) return false;
    }
    return true;
  });
  data.sort(function(a, b) { return String(a[iFecha]) < String(b[iFecha]) ? -1 : 1; });
  return {
    meses: data.map(function(r) { return String(r[iMes]); }),
    flujo: data.map(function(r) { return Number(r[iFlujo]); })
  };
}

function readServicios(ss) {
  var rows = ss.getSheetByName('Servicios').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { name: String(r[0]), color: String(r[1]), ingresos: String(r[2]),
             margen: Number(r[3]), meta: Number(r[4]) };
  });
}

function readFunnel(ss) {
  var rows = ss.getSheetByName('Funnel').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { label: String(r[0]), val: Number(r[1]), pct: Number(r[2]), color: String(r[3]) };
  });
}

function readAlertas(ss) {
  var rows = ss.getSheetByName('Alertas').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { type: String(r[0]), icon: String(r[1]), title: String(r[2]), desc: String(r[3]) };
  });
}

function readDonut(ss) {
  var rows = ss.getSheetByName('DonutServicios').getDataRange().getValues();
  var data = rows.slice(1);
  return {
    labels: data.map(function(r) { return String(r[0]); }),
    data:   data.map(function(r) { return Number(r[1]); }),
    colors: data.map(function(r) { return String(r[2]); })
  };
}

function readPaisesOrigen(ss, fechaInicio, fechaFin, sucursal) {
  var hoja = ss.getSheetByName('PaisesOrigen');
  if (!hoja) return { labels: [], data: [], colors: [] };
  var rows = hoja.getDataRange().getValues();
  if (rows.length < 2) return { labels: [], data: [], colors: [] };
  var hdrs = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iSuc   = hdrs.indexOf('sucursal');
  var iFecha = hdrs.indexOf('fecha');       if (iFecha < 0) iFecha = 1;
  var iPais  = hdrs.indexOf('pais');        if (iPais  < 0) iPais  = 2;
  var iPct   = hdrs.indexOf('porcentaje'); if (iPct   < 0) iPct   = 3;
  var iColor = hdrs.indexOf('color');      if (iColor < 0) iColor = 4;
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[iFecha]).trim();
    if (f < fechaInicio || f > fechaFin) return false;
    if (sucursal && sucursal !== 'Todas' && iSuc >= 0) {
      var s = String(r[iSuc] || '').trim();
      if (s && s !== sucursal) return false;
    }
    return true;
  });
  return {
    labels: data.map(function(r) { return String(r[iPais]); }),
    data:   data.map(function(r) { return Number(r[iPct]); }),
    colors: data.map(function(r) { return String(r[iColor]); })
  };
}

function readCostos(ss) {
  var rows = ss.getSheetByName('DistribucionCostos').getDataRange().getValues();
  var data = rows.slice(1);
  return {
    labels: data.map(function(r) { return String(r[0]); }),
    data:   data.map(function(r) { return Number(r[1]); }),
    colors: data.map(function(r) { return String(r[2]); })
  };
}

/* ══════════════════════════════════════════════════════════════
   ESTADO DE RESULTADOS
   ══════════════════════════════════════════════════════════════ */
function readEstadoResultados(fechaInicio, fechaFin) {
  try {
    var ssEr = SpreadsheetApp.openById(ER_SS_ID);
    var sh = null, allSheets = ssEr.getSheets();
    for (var i = 0; i < allSheets.length; i++) {
      if (allSheets[i].getSheetId() === ER_GID) { sh = allSheets[i]; break; }
    }
    if (!sh) sh = allSheets[0];

    var raw = sh.getDataRange().getValues();
    var tz  = Session.getScriptTimeZone();
    var MON = {jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',
               jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'};

    function parseColDate(val) {
      if (!val && val !== 0) return null;
      if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM');
      var s = String(val).trim();
      var m = s.match(/^([A-Za-z]{3})-(\d{2})$/);
      if (m) return '20' + m[2] + '-' + (MON[m[1].toLowerCase()] || '01');
      return null;
    }

    var headerRow = raw.length > 1 ? raw[1] : [];
    var colMap = [];
    for (var j = 3; j < headerRow.length; j++) {
      var ym = parseColDate(headerRow[j]);
      if (ym) colMap.push({ idx: j, ym: ym });
    }
    var fi = (fechaInicio || '').slice(0, 7);
    var ff = (fechaFin    || '').slice(0, 7);
    var cols = colMap.filter(function(c) {
      return (!fi || c.ym >= fi) && (!ff || c.ym <= ff);
    });

    var METRIC_ROWS = { 'Total Expenses':'total','Gross Profit':'metric','EBITDA':'metric','EBIT':'metric','Net Profit':'metric' };
    var SUBTOTAL_C  = { 'Total Income':true,'Total COGS':true,'Total':true };

    var rows = [];
    for (var i = 4; i < raw.length; i++) {
      var r = raw[i];
      var a = String(r[0]||'').trim(), b = String(r[1]||'').trim(), c = String(r[2]||'').trim();
      if (!a && !b && !c) continue;
      var label, tipo, nivel;
      if      (a && !b && !c) { label=a; tipo=METRIC_ROWS[a]||'header'; nivel=0; }
      else if (!a && b && !c) { label=b; tipo='categoria'; nivel=1; }
      else if (!a && !b && c) { label=c; tipo=SUBTOTAL_C[c]?'subtotal':'dato'; nivel=2; }
      else if (a && b && !c)  { label=b; tipo='subseccion'; nivel=1; }
      else if (!a && b && c)  { label=c==='Total'?b+' total':c; tipo=c==='Total'?'subtotal':'dato'; nivel=2; }
      else                    { label=a||b||c; tipo='dato'; nivel=1; }
      if (!label) continue;
      var valores = cols.map(function(col) {
        var v = r[col.idx]; return (v instanceof Date)?0:(Number(v)||0);
      });
      rows.push({ tipo:tipo, nivel:nivel, label:label, valores:valores,
                  total:valores.reduce(function(s,v){return s+v;},0) });
    }
    return { view:'estado-resultados', fuente:'EstadoResultados',
             meses:cols.map(function(c){return c.ym;}),
             rows:rows, fechaInicio:fechaInicio, fechaFin:fechaFin,
             updated:new Date().toISOString() };
  } catch(ex) {
    return { view:'estado-resultados', fuente:'EstadoResultados',
             error:ex.message, meses:[], rows:[] };
  }
}

/* ══════════════════════════════════════════════════════════════
   OPERATING P&L STATEMENT — 100% stateless, multi-usuario
   ══════════════════════════════════════════════════════════════ */
function readOperatingPL(viewType, plMonth, plYear, plPrevYear) {
  viewType = (viewType || 'Q1').trim();
  plMonth  = (plMonth  || '').trim();
  var thisYear = new Date().getFullYear();
  var yr  = parseInt(plYear     || thisYear, 10) || thisYear;
  var prv = parseInt(plPrevYear || (yr - 1),  10) || (yr - 1);

  var VIEW_OPTIONS = [
    {value:'Q1',label:'Q1 — Primer trimestre'},{value:'Q2',label:'Q2 — Segundo trimestre'},
    {value:'Q3',label:'Q3 — Tercer trimestre'},{value:'Q4',label:'Q4 — Cuarto trimestre'}
  ];
  var MONTH_OPTIONS = [{value:'',label:'— Trimestral —'}];
  ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
   'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
  .forEach(function(m){ MONTH_OPTIONS.push({value:m,label:m}); });
  var yearRange = [];
  for (var yy = yr-3; yy <= yr+3; yy++) yearRange.push(String(yy));

  try {
    var erSS = SpreadsheetApp.openById(ER_SS_ID);
    var sheets = erSS.getSheets();
    var plSheet=null, erSheet=null, bgSheet=null;
    for (var i = 0; i < sheets.length; i++) {
      var gid = sheets[i].getSheetId();
      if (gid===PL_GID)     plSheet=sheets[i];
      if (gid===ER_GID)     erSheet=sheets[i];
      if (gid===BUDGET_GID) bgSheet=sheets[i];
    }
    return _buildPLReport(plSheet, erSheet, bgSheet,
      viewType, plMonth, yr, prv, VIEW_OPTIONS, MONTH_OPTIONS, yearRange);
  } catch(ex) {
    return { view:'p-l', fuente:'OperatingPL', error:ex.message+' L:'+ex.lineNumber,
             rows:[], colHeaders:[], viewOptions:VIEW_OPTIONS, monthOptions:MONTH_OPTIONS };
  }
}

function _buildPLReport(plSheet, erSheet, bgSheet, viewType, plMonth, yr, prv,
                        VIEW_OPTIONS, MONTH_OPTIONS, yearRange) {
  var QTR     = {Q1:[1,2,3],Q2:[4,5,6],Q3:[7,8,9],Q4:[10,11,12]};
  var MES_NUM = {Enero:1,Febrero:2,Marzo:3,Abril:4,Mayo:5,Junio:6,
                 Julio:7,Agosto:8,Septiembre:9,Octubre:10,Noviembre:11,Diciembre:12};
  var targetMonths = plMonth && MES_NUM[plMonth] ? [MES_NUM[plMonth]] : (QTR[viewType]||QTR['Q1']);

  var MON_IDX={jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  function parseColHdr(val) {
    if (!val && val!==0) return null;
    if (val instanceof Date) return {year:val.getFullYear(),month:val.getMonth()+1};
    var s=String(val).trim(), m=s.match(/^([A-Za-z]{3})-(\d{2,4})$/);
    if (!m) return null;
    var mo=MON_IDX[m[1].toLowerCase()]; if(!mo) return null;
    var y=m[2].length===2?2000+parseInt(m[2],10):parseInt(m[2],10);
    return {year:y,month:mo};
  }
  function sumCols(row, cols) {
    var t=0;
    for (var ci=0;ci<cols.length;ci++) {
      var v=row[cols[ci]];
      if (typeof v==='number'){t+=v;continue;}
      if (typeof v==='string'){var n=parseFloat(v.replace(/[$,\s]/g,''));if(!isNaN(n))t+=n;}
    }
    return t;
  }
  function normalize(s){ return String(s||'').trim().toLowerCase(); }

  if (!erSheet) return {view:'p-l',fuente:'OperatingPL',
    error:'Hoja Estado de Resultados no encontrada',
    rows:[],colHeaders:[],viewOptions:VIEW_OPTIONS,monthOptions:MONTH_OPTIONS};

  var erData=erSheet.getDataRange().getValues();
  var erHdr=erData.length>1?erData[1]:[];
  var currCols=[],prevCols=[];
  for (var j=3;j<erHdr.length;j++) {
    var ph=parseColHdr(erHdr[j]); if(!ph) continue;
    if (ph.year===yr  && targetMonths.indexOf(ph.month)>=0) currCols.push(j);
    if (ph.year===prv && targetMonths.indexOf(ph.month)>=0) prevCols.push(j);
  }

  var erLookup={}, totActual=0, totPrev=0;
  for (var ri=4;ri<erData.length;ri++) {
    var r=erData[ri];
    var labA=String(r[0]||'').trim(),labB=String(r[1]||'').trim(),labC=String(r[2]||'').trim();
    if (!labA&&!labB&&!labC) continue;
    var actual=sumCols(r,currCols), prev=sumCols(r,prevCols);
    [labA,labB,labC].forEach(function(lbl){
      if (!lbl) return;
      var k=normalize(lbl);
      if (!erLookup[k]||(actual!==0&&erLookup[k].actual===0)) erLookup[k]={actual:actual,prev:prev};
    });
    if (normalize(labC)==='total income'){totActual=actual;totPrev=prev;}
  }

  var bgLookup={}, totBudget=0;
  if (bgSheet) {
    var bgData=bgSheet.getDataRange().getValues();
    var bgHdr=bgData.length>1?bgData[1]:[];
    var bgCols=[];
    for (var j=3;j<bgHdr.length;j++) {
      var bph=parseColHdr(bgHdr[j]);
      if (bph&&bph.year===yr&&targetMonths.indexOf(bph.month)>=0) bgCols.push(j);
    }
    for (var ri=4;ri<bgData.length;ri++) {
      var br=bgData[ri];
      var ba=String(br[0]||'').trim(),bb=String(br[1]||'').trim(),bc=String(br[2]||'').trim();
      if (!ba&&!bb&&!bc) continue;
      var bv=sumCols(br,bgCols);
      [ba,bb,bc].forEach(function(lbl){
        if (!lbl) return;
        var k=normalize(lbl);
        if (!bgLookup[k]||(bv!==0&&bgLookup[k]===0)) bgLookup[k]=bv;
      });
    }
    totBudget=bgLookup['total income']||0;
  }

  var STRUCTURE=[
    {label:'Revenue',                    tipo:'seccion',indent:0,erAlias:'total income'},
    {label:'Alta',                        tipo:'dato',   indent:1},
    {label:'Surrogacy',                   tipo:'dato',   indent:1},
    {label:'Externos',                    tipo:'dato',   indent:1},
    {label:'Other Income',                tipo:'dato',   indent:1},
    {label:'COGS',                        tipo:'seccion',indent:0,erAlias:'total cogs'},
    {label:'Medicamentos',                tipo:'dato',   indent:1},
    {label:'Laboratorios',                tipo:'dato',   indent:1},
    {label:'Insumos Lab',                 tipo:'dato',   indent:1},
    {label:'Insumos Qx',                  tipo:'dato',   indent:1},
    {label:'Comisiones',                  tipo:'dato',   indent:1},
    {label:'Honorarios',                  tipo:'dato',   indent:1},
    {label:'Honorarios Consultas',        tipo:'dato',   indent:1},
    {label:'Other Costs',                 tipo:'dato',   indent:1},
    {label:'Gross Profit',                tipo:'metric', indent:0},
    {label:'OpEx',                        tipo:'seccion',indent:0,erAlias:'total opex'},
    {label:'Marketing',                   tipo:'dato',   indent:1},
    {label:'Renta',                       tipo:'dato',   indent:1},
    {label:'Mto Renta',                   tipo:'dato',   indent:1},
    {label:'Servicios',                   tipo:'dato',   indent:1},
    {label:'Clinic Contribution',         tipo:'metric', indent:0},
    {label:'G&A',                         tipo:'seccion',indent:0,erAlias:'total g&a'},
    {label:'Nomina',                      tipo:'dato',   indent:1},
    {label:'Gastos Varios',               tipo:'dato',   indent:1},
    {label:'EBITDA',                      tipo:'metric', indent:0},
    {label:'Depreciation & Amortization',tipo:'dato',   indent:1},
    {label:'EBIT',                        tipo:'metric', indent:0},
    {label:'EBT',                         tipo:'metric', indent:0},
    {label:'Taxes',                       tipo:'dato',   indent:1},
    {label:'Net Profit',                  tipo:'metric', indent:0}
  ];

  var rows=[];
  for (var si=0;si<STRUCTURE.length;si++) {
    var s=STRUCTURE[si];
    var key=s.erAlias||normalize(s.label);
    var er=erLookup[key]||{actual:0,prev:0};
    var bg=bgLookup[key]||bgLookup[normalize(s.label)]||0;
    rows.push({
      label:s.label, tipo:s.tipo, indent:s.indent,
      values:[
        er.actual,
        totActual  ? er.actual/totActual  : null,
        bg,
        totBudget  ? bg/totBudget         : null,
        bg         ? (er.actual-bg)/Math.abs(bg)       : null,
        er.prev,
        totPrev    ? er.prev/totPrev      : null,
        er.prev    ? (er.actual-er.prev)/Math.abs(er.prev) : null
      ]
    });
  }

  var pLabel=plMonth||viewType;
  return {
    view:'p-l', fuente:'OperatingPL',
    viewType:viewType, activeMonth:plMonth,
    currentYear:String(yr), currentPrev:String(prv),
    viewOptions:VIEW_OPTIONS, monthOptions:MONTH_OPTIONS,
    hasMonthFilter:true, yearRange:yearRange,
    colHeaders:[
      {label:'Actual '+pLabel+' '+yr,    isPct:false,isVs:false},
      {label:'%',                          isPct:true, isVs:false},
      {label:'Budget '+pLabel+' '+yr,    isPct:false,isVs:false},
      {label:'%',                          isPct:true, isVs:false},
      {label:'vs. Budget',                 isPct:true, isVs:true },
      {label:prv+' '+pLabel,             isPct:false,isVs:false},
      {label:'%',                          isPct:true, isVs:false},
      {label:'YOY %',                      isPct:true, isVs:true }
    ],
    rows:rows,
    _debug:{currCols:currCols,prevCols:prevCols,totActual:totActual,
            hdrSample:erHdr.slice(3,9).map(function(v){
              return (v instanceof Date)?'[Date:'+v.getFullYear()+'-'+(v.getMonth()+1)+']':String(v);
            })}
  };
}

/* ══════════════════════════════════════════════════════════════
   BANCOS — 3 cuentas + escritura por pestaña
   ══════════════════════════════════════════════════════════════ */
function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents;
    if (!raw) return jsonResponse({error:'Sin datos POST'});
    var body = JSON.parse(raw);
    if (body.action === 'saveBankRow') return jsonResponse(saveBankRow(body.banco, body.row));
    return jsonResponse({error:'Accion desconocida: ' + body.action});
  } catch(ex) { return jsonResponse({error: ex.message}); }
}

function readBanksData() {
  try {
    var ss = SpreadsheetApp.openById(BANKS_SS_ID);
    var sh = ss.getSheets();
    function byGid(gid){ for(var i=0;i<sh.length;i++) if(sh[i].getSheetId()===gid) return sh[i]; return null; }
    function num(v){ if(typeof v==='number') return v; var n=parseFloat(String(v||'').replace(/[$,\s]/g,'')); return isNaN(n)?0:n; }
    function dt(v){ if(!v) return ''; if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0'); return String(v); }

    function rSant(sheet) {
      var B={id:'santander',nombre:'Santander',color:'#ec0000',saldo:0,movimientos:[],totalRows:0};
      if(!sheet) return B;
      var r=sheet.getDataRange().getValues(); if(r.length<2) return B;
      for(var i=r.length-1;i>=1;i--){ var s=num(r[i][3]); if(s!==0){B.saldo=s;break;} }
      B.totalRows=r.length-1;
      B.movimientos=r.slice(1).slice(-30).reverse().map(function(x){
        var d=num(x[1]),t=num(x[2]);
        return{fecha:dt(x[0]),deposito:d,retiro:t,monto:d>0?d:-t,saldo:num(x[3]),
               referencia:String(x[4]||''),depositoUSD:num(x[5]),tipoCambio:num(x[6]),
               poliza:String(x[7]||''),observaciones:String(x[8]||''),tipo:d>0?'deposito':'retiro'};
      });
      return B;
    }
    function rAmex(sheet) {
      var B={id:'amex',nombre:'AMEX',color:'#007bc1',saldo:0,movimientos:[],totalRows:0};
      if(!sheet) return B;
      var r=sheet.getDataRange().getValues(); if(r.length<2) return B;
      for(var i=r.length-1;i>=1;i--){ var s=num(r[i][2]); if(s!==0){B.saldo=s;break;} }
      B.totalRows=r.length-1;
      B.movimientos=r.slice(1).slice(-30).reverse().map(function(x){
        var m=num(x[1]);
        return{fecha:dt(x[0]),monto:m,saldo:num(x[2]),referencia:String(x[3]||''),
               usd:num(x[4]),tipoCambio:num(x[5]),notas:String(x[6]||''),
               poliza:String(x[7]||''),mes:String(x[8]||''),tipo:m>=0?'cargo':'pago'};
      });
      return B;
    }
    function rMP(sheet) {
      var B={id:'mercadopago',nombre:'Mercado Pago',color:'#009ee3',saldo:0,movimientos:[],totalRows:0};
      if(!sheet) return B;
      var r=sheet.getDataRange().getValues(); if(r.length<2) return B;
      for(var i=r.length-1;i>=1;i--){ var s=num(r[i][6]); if(s!==0){B.saldo=s;break;} }
      B.totalRows=r.length-1;
      B.movimientos=r.slice(1).slice(-30).reverse().map(function(x){
        return{mes:String(x[0]||''),fecha:dt(x[1]),cobro:num(x[2]),comisiones:num(x[3]),
               pctComision:num(x[4]),totalVenta:num(x[5]),saldo:num(x[6]),
               liberado:x[7]===true||String(x[7]).toUpperCase()==='TRUE',
               observaciones:String(x[8]||''),tipo:String(x[9]||'CARGO').toUpperCase(),monto:num(x[5])};
      });
      return B;
    }

    var sant=rSant(byGid(BANKS_GID.santander));
    var amex=rAmex(byGid(BANKS_GID.amex));
    var mp  =rMP(byGid(BANKS_GID.mercadopago));
    return { view:'cashflow', fuente:'CashFlow',
             bancos:{santander:sant,amex:amex,mercadopago:mp},
             totalSaldo:sant.saldo+amex.saldo+mp.saldo };
  } catch(ex) {
    return { view:'cashflow', fuente:'CashFlow', error:ex.message, bancos:{}, totalSaldo:0 };
  }
}

function saveBankRow(banco, row) {
  try {
    var ss=SpreadsheetApp.openById(BANKS_SS_ID);
    var sh=ss.getSheets();
    var key=String(banco).toLowerCase().replace(/[\s-]/g,'');
    var gid=BANKS_GID[key];
    if (gid===undefined) return {ok:false, error:'Banco desconocido: '+banco};
    var sheet=null;
    for(var i=0;i<sh.length;i++) if(sh[i].getSheetId()===gid){sheet=sh[i];break;}
    if(!sheet) return {ok:false, error:'Pestaña no encontrada'};
    // Columna del saldo por banco: Santander=col4, AMEX=col3, MP=col7
    var sc=(key==='santander')?4:(key==='amex')?3:7;
    var lr=sheet.getLastRow(), ls=0;
    if(lr>1){
      var sv=sheet.getRange(lr,sc).getValue();
      ls=(typeof sv==='number')?sv:parseFloat(String(sv).replace(/[$,]/g,''))||0;
    }
    if      (key==='santander') row[3]=ls+(parseFloat(row[1])||0)-(parseFloat(row[2])||0);
    else if (key==='amex')      row[2]=ls+(parseFloat(row[1])||0);
    else                        row[6]=ls+(parseFloat(row[5])||0);
    sheet.appendRow(row);
    return {ok:true, banco:banco, newSaldo:row[sc-1], totalRows:sheet.getLastRow()-1};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

/* ── Constantes definidas en api_config.gs (mismo proyecto GAS) ──
   SHEET_ID, AUTH_SECRET, ER_SS_ID, BANKS_SS_ID, LAB_SS_ID, MED_SS_ID,
   PAC_SS_ID, PROD_SS_ID, QX_SS_ID, ER_GID, BUDGET_GID, PL_GID,
   BANKS_GID, CAPTURA_SHEETS, SHEET_ALIASES, CAPTURA_SHEET_ID_DEFAULT
   ────────────────────────────────────────────────────────────────── */

/* ── Autenticación: helpers ──────────────────────────────────── */
function sha256Hex(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str)
    .map(function(b){ return ('0'+(b&0xFF).toString(16)).slice(-2); }).join('');
}
function generateToken(email, day) {
  day = day || fmtDate(new Date());
  return Utilities.base64Encode(email+'|'+day+'|'+sha256Hex(email+'|'+day+'|'+AUTH_SECRET));
}
function verifyToken(token) {
  if (!token) return null;
  try {
    var dec   = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    var parts = dec.split('|');
    if (parts.length < 3) return null;
    var email = parts[0], day = parts[1];
    var diff  = (new Date() - new Date(day)) / 86400000;
    if (diff > 7 || diff < -1) return null;
    return (generateToken(email, day) === token) ? email : null;
  } catch(ex){ return null; }
}
function getUserRow(ss, email) {
  var sh = ss.getSheetByName('Usuarios');
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  var h    = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
  var eI=h.indexOf('email'), nI=h.indexOf('nombre'), pI=h.indexOf('contraseña'),
      rI=h.indexOf('rol'),   aI=h.indexOf('activo');
  for (var i=1; i<data.length; i++) {
    if (String(data[i][eI]).trim().toLowerCase() === email.toLowerCase()) {
      return { email: String(data[i][eI]).trim(), nombre: String(data[i][nI>-1?nI:0]).trim(),
               password: pI>-1?String(data[i][pI]).trim():'',
               rol: rI>-1?String(data[i][rI]).trim():'viewer',
               activo: aI>-1?data[i][aI]:true, rowNum: i+1 };
    }
  }
  return null;
}
function getRolConfig(ss, rol) {
  var sh  = ss.getSheetByName('Roles');
  var def = { vistasBloqueadas:[], soloLectura:false };
  if (!sh) return def;
  var data = sh.getDataRange().getValues();
  var h    = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
  var rI=h.indexOf('rol'), bI=h.indexOf('vistas_bloqueadas'), lI=h.indexOf('solo_lectura');
  for (var i=1; i<data.length; i++) {
    if (String(data[i][rI]).trim().toLowerCase() === rol.toLowerCase()) {
      var bloq = bI>-1 ? String(data[i][bI]).split(',').map(function(v){return v.trim();}).filter(Boolean) : [];
      return { vistasBloqueadas: bloq, soloLectura: lI>-1 ? !!data[i][lI] : false };
    }
  }
  return def;
}


/* doGet fue movido a core.gs — es el enrutador principal del ERP.
   No definir doGet aquí para evitar que sobreescriba al de core.gs. */

/* Borra todas las entradas de caché relacionadas con una hoja (fuente) */
function invalidateViewCache(sheetName) {
  try {
    var cache = CacheService.getScriptCache();
    // La clave incluye el viewId que puede ser el nombre de hoja directamente.
    // Borramos con el nombre exacto y variantes comunes.
    var keys = [];
    // Generar keys para los últimos 12 meses como rango de fechas posible
    var now = new Date();
    for (var i = -1; i <= 12; i++) {
      var d1 = new Date(now.getFullYear(), now.getMonth() - 6 + i, 1);
      var d2 = new Date(now.getFullYear(), now.getMonth() - 6 + i + 7, 0);
      keys.push('v2_' + sheetName + '_' + fmtDate(d1) + '_' + fmtDate(d2));
    }
    cache.removeAll(keys);
  } catch(ignored) {}
}


/* ══════════════════════════════════════════════════════════════
   LEE LA HOJA Menu
   Columnas: ID | Padre | Label | Seccion | Icono | Orden | Tipo | Fuente | Activo
   ══════════════════════════════════════════════════════════════ */
function readMenu(ss) {
  var rows = ss.getSheetByName('Menu').getDataRange().getValues();
  return rows.slice(1)
    .filter(function(r) {
      // Solo filas activas (columna I = TRUE o vacío)
      var activo = r[8];
      return activo !== false && String(activo).toUpperCase() !== 'FALSE';
    })
    .map(function(r) {
      return {
        id:      String(r[0]).trim(),
        padre:   String(r[1]).trim(),
        label:   String(r[2]).trim(),
        seccion: String(r[3]).trim(),
        icono:   String(r[4]).trim() || 'circle',
        orden:   Number(r[5]) || 0,
        tipo:    String(r[6]).trim().toLowerCase(), // 'vista' | 'grupo'
        fuente:  String(r[7]).trim(),               // 'mensual' | NombreHoja
        activo:  r[8] !== false
      };
    });
}

/* ══════════════════════════════════════════════════════════════
   ENRUTADOR DE DATOS POR VISTA
   fuente = 'mensual' → datos financieros filtrados por rango de fechas
   fuente = NombreHoja → lee esa hoja como tabla de captura (sin filtro)
   ══════════════════════════════════════════════════════════════ */
function _plMatch(s) {
  var n = (s||'').toLowerCase().replace(/[\s&_\-]+/g,'');
  // Coincide con: pl, fin-pl, operatingpl, p&l, pyg, etc.
  return n === 'pl' || n === 'operatingpl' || n === 'pyg' ||
         n === 'finpl' || n.indexOf('operatingpl') >= 0 ||
         /^.{0,6}pl$/.test(n); // termina en "pl" con hasta 6 chars de prefijo
}

function readViewData(ss, viewId, fechaInicio, fechaFin, sucursal, viewType, plMonth, plYear, plPrevYear) {
  sucursal   = sucursal   || 'Todas';
  viewType   = viewType   || '';
  plMonth    = plMonth    || '';
  plYear     = plYear     || '';
  plPrevYear = plPrevYear || '';
  var menu = readMenu(ss);
  var item = null;
  for (var i = 0; i < menu.length; i++) {
    if (menu[i].id === viewId) { item = menu[i]; break; }
  }
  var fuente = item ? item.fuente : 'mensual';
  // Traducir alias (ej: Orden_Compra → Ent. Med)
  if (fuente && SHEET_ALIASES[fuente]) fuente = SHEET_ALIASES[fuente];

  // Vistas de configuración local — no leen Sheets
  if (/^(configuracion|ajustes|config)$/i.test(viewId) || /^(configuracion|ajustes|config)$/i.test(fuente||'')) {
    return { viewId: viewId, fuente: fuente, rows: [], headers: [], _isConfig: true };
  }

  // P&L check ANTES del fallback mensual — el fuente puede estar vacío en el menú
  if (_plMatch(fuente) || _plMatch(viewId)) {
    return readOperatingPL(viewType, plMonth, plYear, plPrevYear);
  }

  if (!fuente || fuente === 'mensual') {
    return readMensualData(ss, fechaInicio, fechaFin, viewId, sucursal);
  }

  if (fuente === 'med-dashboard' || viewId === 'med-resumen' || viewId === 'prod-medicamentos') {
    return readMedDashboard(ss, fechaInicio, fechaFin);
  }

  if (fuente === 'lab-resumen' || viewId === 'lab-resumen' || viewId === 'lab-resumen-dash') {
    return readLabResumen(fechaInicio, fechaFin);
  }

  if (fuente === 'qx-resumen' || viewId === 'qx-resumen') {
    return readQxResumen(fechaInicio, fechaFin);
  }

  if (viewId === 'gestion-roles') {
    return readGestionRoles(ss);
  }

  function _erMatch(s) {
    var n = (s || '').toLowerCase().replace(/\bde\b/g,'').replace(/[\s_\-]+/g,'');
    return n.indexOf('estado') > -1 && n.indexOf('result') > -1;
  }
  if (_erMatch(fuente) || _erMatch(viewId)) {
    return readEstadoResultados(fechaInicio, fechaFin);
  }

  if (fuente === 'Rep Ejecutivo' || viewId === 'rep-ejecutivo') {
    return readRepEjecutivo(ss, fechaInicio, fechaFin);
  }

  // Vistas de Reportes: devolver placeholder vacío hasta que se creen las hojas
  if (/^rep-/.test(viewId)) {
    return { view: viewId, fuente: fuente, rows: [], headers: [], periodo: fechaInicio + ' — ' + fechaFin };
  }

  // Cash Flow / Flujo de Efectivo — los datos reales vienen de action=banks
  if (/flujo|cashflow|cash.flow|efectivo|fin.cf/i.test(viewId) ||
      /flujo|cashflow|cash.flow|efectivo/i.test(fuente||'')) {
    return readBanksData();
  }

  return readCapturaData(ss, fuente, viewId, fechaInicio, fechaFin, sucursal);
}

/* ══════════════════════════════════════════════════════════════
   RESUMEN LABORATORIO — Dashboard clínico de calidad embrionaria
   Fuente: spreadsheet Lab (LAB_SS_ID)
   Hojas leídas: Resumen, ART Lab, FET, Inventario Crío, Insumos
   ══════════════════════════════════════════════════════════════ */
function readLabResumen(fechaInicio, fechaFin) {

  // ── CONFIGURACIÓN DE COLUMNAS ──────────────────────────────────
  // ART Lab: encabezados en fila 1 + fila 2 fusionadas (A=0, B=1…)
  // Ajusta estos índices si agregas/mueves columnas en la hoja
  var ART_COL_MES       = 0;   // A: Mes-Año
  var ART_COL_FECHA     = 1;   // B: Date
  var ART_COL_OOCITOS   = 9;   // J: # oocytes (recuperados)
  // Cuando conozcas las columnas de 2PN y blastocistos, actualiza:
  var ART_COL_2PN       = -1;  // -1 = no configurado aún
  var ART_COL_BLASTO    = -1;  // -1 = no configurado aún
  var ART_COL_ICSI_DANO = -1;  // -1 = no configurado aún

  // FET: encabezado en fila 2 (fila 1 vacía)
  var FET_COL_FECHA    = 0;  // A: Fecha
  var FET_COL_SURVIVED = 5;  // F: Survived ("Si"/"No")
  var FET_COL_BETA     = 6;  // G: Beta
  var FET_COL_PREG     = 7;  // H: Clinical Preg.

  // Inventario Crío: encabezado en fila 1
  var CRIO_COL_NOMBRE  = 0;  // A: Nombre paciente
  var CRIO_COL_FECHA   = 1;  // B: Fecha Crío
  var CRIO_COL_OOV     = 2;  // C: Oov (ovocitos)
  var CRIO_COL_EMB     = 3;  // D: Emb (embriones)

  // Insumos: encabezado en fila 2 (fila 1 vacía); col A vacía → datos desde col B
  var INS_COL_INSUMO   = 3;  // D: Insumo
  var INS_COL_PROV     = 4;  // E: Proveedor
  var INS_COL_FECHA    = 2;  // C: Fecha
  var INS_COL_COSTO    = 8;  // I: Costo
  // ─────────────────────────────────────────────────────────────

  function pct(num, den) {
    if (!den || den === 0) return null;
    return Math.round((num / den) * 1000) / 10;
  }
  function fmtFecha(v) {
    if (!v) return '';
    if (v instanceof Date) return fmtDate(v);
    var d = new Date(v);
    return isNaN(d.getTime()) ? String(v) : fmtDate(d);
  }
  function mesLabel(v) {
    if (!v) return '';
    if (v instanceof Date) return (v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0'));
    var s = String(v);
    var m = s.match(/(\d{4})-(\d{2})/);
    return m ? m[0] : s.slice(0,7);
  }

  try {
    var ssLab = SpreadsheetApp.openById(LAB_SS_ID);

    // ── 1. Leer ART Lab (encabezados fusionados fila 1+2, datos desde fila 3) ──
    var shArt  = findSheet(ssLab, 'ART Lab');
    var artRaw = shArt ? shArt.getDataRange().getValues() : [];
    // Fusionar headers de fila 1 y fila 2
    var artData = artRaw.slice(2).filter(function(r) {
      return r.some(function(c) { return String(c).trim() !== ''; });
    });
    var totalCiclos = artData.length;
    var totalOocitos = artData.reduce(function(s,r) { return s + (Number(r[ART_COL_OOCITOS])||0); }, 0);

    // % Fecundación y Blastocistos: calcular si las columnas están configuradas
    var total2PN   = ART_COL_2PN   > -1 ? artData.reduce(function(s,r){ return s+(Number(r[ART_COL_2PN])||0); },0) : null;
    var totalBlasto= ART_COL_BLASTO > -1 ? artData.reduce(function(s,r){ return s+(Number(r[ART_COL_BLASTO])||0); },0) : null;

    // Tendencia por mes desde ART Lab
    var artPorMes = {};
    artData.forEach(function(r) {
      var mes = mesLabel(r[ART_COL_FECHA] || r[ART_COL_MES]);
      if (!mes) return;
      if (!artPorMes[mes]) artPorMes[mes] = { ciclos:0, oocitos:0, pn2:0, blasto:0 };
      artPorMes[mes].ciclos++;
      artPorMes[mes].oocitos += Number(r[ART_COL_OOCITOS]) || 0;
      if (ART_COL_2PN    > -1) artPorMes[mes].pn2    += Number(r[ART_COL_2PN])    || 0;
      if (ART_COL_BLASTO > -1) artPorMes[mes].blasto += Number(r[ART_COL_BLASTO]) || 0;
    });
    var mesesArt = Object.keys(artPorMes).sort().slice(-6);

    // ── 2. Leer FET (encabezado en fila 2, datos desde fila 3) ──
    var shFet  = findSheet(ssLab, 'FET');
    var fetRaw = shFet ? shFet.getDataRange().getValues() : [];
    var fetData = fetRaw.slice(2).filter(function(r) {
      return r.some(function(c) { return String(c).trim() !== ''; });
    });
    var totalFet      = fetData.length;
    var fetSurvividos = fetData.filter(function(r) {
      return String(r[FET_COL_SURVIVED]).trim().toLowerCase() === 'si';
    }).length;
    var fetPregnancy  = fetData.filter(function(r) {
      return String(r[FET_COL_PREG]).trim().toLowerCase() === 'si';
    }).length;

    // Tendencia FET por mes
    var fetPorMes = {};
    fetData.forEach(function(r) {
      var mes = mesLabel(r[FET_COL_FECHA]);
      if (!mes) return;
      if (!fetPorMes[mes]) fetPorMes[mes] = { total:0, survived:0 };
      fetPorMes[mes].total++;
      if (String(r[FET_COL_SURVIVED]).trim().toLowerCase() === 'si') fetPorMes[mes].survived++;
    });

    // ── 3. Inventario Crío (encabezado fila 1, datos desde fila 2) ──
    var shCrio  = findSheet(ssLab, 'Inventario Crío');
    var crioRaw = shCrio ? shCrio.getDataRange().getValues() : [];
    var crioData = crioRaw.slice(1).filter(function(r) {
      return String(r[CRIO_COL_NOMBRE]).trim() !== '';
    });
    var totalOvCrio  = crioData.reduce(function(s,r){ return s+(Number(r[CRIO_COL_OOV])||0); },0);
    var totalEmbCrio = crioData.reduce(function(s,r){ return s+(Number(r[CRIO_COL_EMB])||0); },0);

    // ── 4. Insumos (encabezado fila 2, datos desde fila 3) ──
    var shIns  = findSheet(ssLab, 'Insumos');
    var insRaw = shIns ? shIns.getDataRange().getValues() : [];
    var insData = insRaw.slice(2).filter(function(r) {
      return String(r[INS_COL_INSUMO]).trim() !== '';
    });

    // ── Construir respuesta ──
    var fetPct = pct(fetSurvividos, totalFet);
    var fecPct = (ART_COL_2PN > -1) ? pct(total2PN, totalOocitos) : null;
    var blaPct = (ART_COL_BLASTO > -1) ? pct(totalBlasto, total2PN) : null;

    return {
      view:   'lab-resumen',
      fuente: 'lab-resumen',
      kpis: {
        fecundacion:      fecPct,
        blastocistos:     blaPct,
        fetSupervivencia: fetPct,
        icsiDano:         null,           // configurar ART_COL_ICSI_DANO
        totalCiclos:      totalCiclos,
        totalFet:         totalFet,
        fetPregnancy:     fetPregnancy,
        totalOvCrio:      totalOvCrio,
        totalEmbCrio:     totalEmbCrio,
        mesPeriodo:       mesesArt.length ? mesesArt[mesesArt.length-1] : ''
      },
      tendencia: {
        meses:        mesesArt,
        ciclos:       mesesArt.map(function(m){ return (artPorMes[m]||{}).ciclos||0; }),
        fecundacion:  mesesArt.map(function(m){
          var d = artPorMes[m]||{}; return ART_COL_2PN>-1 ? pct(d.pn2||0, d.oocitos||0) : null;
        }),
        fetSupervivencia: mesesArt.map(function(m){
          var d = fetPorMes[m]||{}; return pct(d.survived||0, d.total||0);
        })
      },
      tanques: crioData.slice(0,10).map(function(r) {
        return {
          nombre: String(r[CRIO_COL_NOMBRE]).split(' ').slice(0,2).join(' '),
          oocitos: Number(r[CRIO_COL_OOV])||0,
          embriones: Number(r[CRIO_COL_EMB])||0
        };
      }),
      insumos: insData.map(function(r) {
        return {
          item:        String(r[INS_COL_INSUMO]),
          proveedor:   String(r[INS_COL_PROV] || ''),
          fecha:       fmtFecha(r[INS_COL_FECHA]),
          costo:       Number(r[INS_COL_COSTO]) || 0,
          estado:      'ok'
        };
      })
    };
  } catch(ex) {
    return { view: 'lab-resumen', fuente: 'lab-resumen', error: ex.message,
             kpis: {}, tendencia: { meses:[], ciclos:[], fecundacion:[], fetSupervivencia:[] },
             tanques: [], insumos: [] };
  }
}

/* ══ DASHBOARD QUIROFANO ════════════════════════════════════════
   Placeholder — ampliar cuando se cree el spreadsheet de Qx
   ══════════════════════════════════════════════════════════════ */
function readQxResumen(fechaInicio, fechaFin) {
  try {
    var ssQx = SpreadsheetApp.openById(QX_SS_ID);
    // Intentar leer hoja "Insumos Qx" para conteos básicos
    var insSheet = findSheet(ssQx, 'Insumos Qx');
    var totalInsumos = 0, totalCosto = 0;
    if (insSheet) {
      var allRows = insSheet.getDataRange().getValues();
      var hdrs = allRows[0];
      var colCosto = hdrs.indexOf('Costo');
      allRows.slice(1).forEach(function(r) {
        if (r[0]) {
          totalInsumos++;
          if (colCosto >= 0 && r[colCosto]) totalCosto += parseFloat(r[colCosto]) || 0;
        }
      });
    }
    return {
      view: 'qx-resumen', fuente: 'qx-resumen',
      kpis: [
        { label: 'Insumos registrados', value: totalInsumos, format: 'number', icon: 'package' },
        { label: 'Costo total insumos', value: totalCosto,   format: 'currency', icon: 'dollar-sign' }
      ],
      rows: [], headers: []
    };
  } catch(ex) {
    return { view: 'qx-resumen', fuente: 'qx-resumen', kpis: [], rows: [], headers: [], error: ex.message };
  }
}

/* ══ REPORTE EJECUTIVO — Agrega KPIs de todas las secciones ═══ */
function readRepEjecutivo(ss, fechaInicio, fechaFin) {
  var result = {
    view: 'rep-ejecutivo', fuente: 'Rep Ejecutivo',
    periodo: (fechaInicio || '') + ' — ' + (fechaFin || ''),
    sections: { financiero: {}, clinico: {}, operaciones: {} },
    rows: [], headers: []
  };

  // ── FINANCIERO: leer Mensual_Todos filtrado por período ──────
  try {
    var mensualSheet = findSheet(ss, 'Mensual_Todos');
    if (mensualSheet) {
      var mRows = mensualSheet.getDataRange().getValues();
      var mHdrs = mRows[0];
      var colFecha    = mHdrs.indexOf('Fecha');
      var colIngresos = mHdrs.indexOf('Ingresos');
      var colGastos   = mHdrs.indexOf('Gastos');
      var totIng = 0, totGas = 0, found = false;
      mRows.slice(1).forEach(function(r) {
        var f = colFecha >= 0 ? String(r[colFecha]).slice(0, 10) : '';
        if (f && fechaInicio && f < fechaInicio) return;
        if (f && fechaFin   && f > fechaFin)   return;
        if (colIngresos >= 0) totIng += parseFloat(r[colIngresos]) || 0;
        if (colGastos   >= 0) totGas += parseFloat(r[colGastos])   || 0;
        found = true;
      });
      if (found || totIng || totGas) {
        result.sections.financiero.ingresos = totIng;
        result.sections.financiero.gastos   = totGas;
      }
    }
  } catch(e) {}

  // ── CLÍNICO: leer Lab SS ──────────────────────────────────────
  try {
    var ssLab = SpreadsheetApp.openById(LAB_SS_ID);

    var artSheet = findSheet(ssLab, 'ART Lab');
    if (artSheet) {
      var artRows = artSheet.getDataRange().getValues().slice(2);
      result.sections.clinico.ciclosART = artRows.filter(function(r){ return r[0]; }).length;
    }

    var fetSheet = findSheet(ssLab, 'FET');
    if (fetSheet) {
      var fetRows = fetSheet.getDataRange().getValues().slice(2);
      result.sections.clinico.fetRealizados = fetRows.filter(function(r){ return r[0]; }).length;
    }

    // Banco Crío
    var crioSheet = findSheet(ssLab, 'Inventario Crío') || findSheet(ssLab, 'Inventario Crio');
    if (crioSheet) {
      var crioRows = crioSheet.getDataRange().getValues().slice(1);
      result.sections.operaciones.bancoCrio = crioRows.filter(function(r){ return r[0]; }).length;
    }

    // Insumos activos (Lab)
    var insSheet = findSheet(ssLab, 'Insumos');
    if (insSheet) {
      var insRows = insSheet.getDataRange().getValues().slice(1);
      result.sections.operaciones.insumosActivos = insRows.filter(function(r){ return r[0]; }).length;
    }
  } catch(e) {}

  // ── OPERACIONES: alertas de inventario desde hoja Alertas ────
  try {
    var alertSheet = findSheet(ss, 'Alertas');
    if (alertSheet) {
      var alertRows = alertSheet.getDataRange().getValues().slice(1);
      result.sections.operaciones.alertasInventario = alertRows.filter(function(r){ return r[0]; }).length;
    }
  } catch(e) {}

  return result;
}

/* ══ DASHBOARD MEDICAMENTOS ════════════════════════════════════ */
function readMedDashboard(ss, fechaInicio, fechaFin) {
  var medId = CAPTURA_SHEETS['Medicamentos'] || CAPTURA_SHEET_ID_DEFAULT;
  var ssMed = SpreadsheetApp.openById(medId);

  function readSheet(nombre) {
    var sh = findSheet(ssMed, nombre);
    if (!sh || sh.getLastRow() < 2) return { headers: [], rows: [] };
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var rows = vals.slice(1).filter(function(r){
      return r.some(function(c){ return String(c).trim() !== ''; });
    }).map(function(r){
      var obj = {};
      hdrs.forEach(function(h, i){ obj[h] = r[i]; });
      return obj;
    });
    // Filtrar por fecha si existe columna Fecha
    var tieneFecha = hdrs.indexOf('Fecha') !== -1;
    if (tieneFecha && fechaInicio && fechaFin) {
      rows = rows.filter(function(r){
        var raw = r['Fecha'];
        var f = (raw instanceof Date)
          ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(raw || '').slice(0, 10);
        return f >= fechaInicio && f <= fechaFin;
      });
    }
    return { headers: hdrs, rows: rows };
  }

  var compraData = readSheet('Ent. Med');
  var estimData  = readSheet('Estimulación');
  var compras    = compraData.rows;
  var estims     = estimData.rows;

  // ── Agregaciones compras ──────────────────────────────────────
  var comprasPorMed = {}, gastosPorMed = {}, comprasPorMes = {};
  compras.forEach(function(r) {
    var med = String(r['Medicamento'] || '').trim();
    var cant = Number(r['Cantidad']) || 0;
    var total = Number(r['Total']) || (cant * (Number(r['Precio_Unitario']) || 0));
    var mes = String(r['Fecha'] || '').slice(0, 7);
    if (med) {
      comprasPorMed[med] = (comprasPorMed[med] || 0) + cant;
      gastosPorMed[med]  = (gastosPorMed[med]  || 0) + total;
    }
    if (mes) comprasPorMes[mes] = (comprasPorMes[mes] || 0) + cant;
  });

  // ── Columnas de medicamentos: todo lo que viene DESPUÉS de "Cancelado" ──
  // Al agregar nuevas columnas en Sheets después de Cancelado se incluyen automáticamente
  var estimHeaders = estimData.headers || [];
  var canceladoIdx = -1;
  for (var ci = 0; ci < estimHeaders.length; ci++) {
    if (estimHeaders[ci].trim().toLowerCase() === 'cancelado') { canceladoIdx = ci; break; }
  }
  var MED_COLS = canceladoIdx >= 0
    ? estimHeaders.slice(canceladoIdx + 1).filter(function(h){ return h.trim() !== ''; })
    : [];
  var usosPorMed = {}, usosPorMes = {}, pacientesSet = {};
  estims.forEach(function(r) {
    var pac = String(r['Paciente'] || '').trim();
    var mes = String(r['Fecha']    || '').slice(0, 7);
    if (pac) pacientesSet[pac] = 1;
    var totalFila = 0;
    MED_COLS.forEach(function(col) {
      var cant = Number(r[col]) || 0;
      if (cant > 0) {
        usosPorMed[col] = (usosPorMed[col] || 0) + cant;
        totalFila += cant;
      }
    });
    if (mes && totalFila > 0) usosPorMes[mes] = (usosPorMes[mes] || 0) + totalFila;
  });

  // ── Top 8 ─────────────────────────────────────────────────────
  function top8(obj) {
    return Object.keys(obj).map(function(k){ return { label: k, value: obj[k] }; })
      .sort(function(a,b){ return b.value - a.value; }).slice(0, 8);
  }

  var topCompras = top8(comprasPorMed);
  var topUsos    = top8(usosPorMed);

  // ── Evolución mensual ─────────────────────────────────────────
  var mesesSet = {};
  Object.keys(comprasPorMes).forEach(function(m){ mesesSet[m]=1; });
  Object.keys(usosPorMes).forEach(function(m){ mesesSet[m]=1; });
  var meses = Object.keys(mesesSet).sort();

  // ── KPIs ──────────────────────────────────────────────────────
  var totalCompras = compras.reduce(function(s,r){ return s+(Number(r['Cantidad'])||0); }, 0);
  var totalUsos    = Object.values ? Object.values(usosPorMed).reduce(function(s,v){ return s+v; }, 0)
                    : Object.keys(usosPorMed).reduce(function(s,k){ return s+usosPorMed[k]; }, 0);
  var gastoTotal   = estims.reduce(function(s,r){ return s+(Number(r['Costo Meds'])||0); }, 0)
                    + compras.reduce(function(s,r){
    return s + (Number(r['Total']) || (Number(r['Cantidad'])||0)*(Number(r['Precio_Unitario'])||0));
  }, 0);

  return {
    view:   'med-resumen',
    periodo: fechaInicio.slice(0,7) + ' → ' + fechaFin.slice(0,7),
    kpis: {
      totalCompras:       totalCompras,
      totalUsos:          totalUsos,
      gastoTotal:         gastoTotal,
      medsDistintos:      Object.keys(comprasPorMed).length,
      pacientesAtendidos: Object.keys(pacientesSet).length
    },
    topCompras: topCompras,
    topUsos:    topUsos,
    evolucion: {
      meses:   meses,
      compras: meses.map(function(m){ return comprasPorMes[m] || 0; }),
      usos:    meses.map(function(m){ return usosPorMes[m]    || 0; })
    }
  };
}

/* ── Datos financieros completos (filtrados por rango de fechas) ─── */
function readMensualData(ss, fechaInicio, fechaFin, viewId, sucursal) {
  sucursal = sucursal || 'Todas';
  var label = fechaInicio.slice(0,7) + ' → ' + fechaFin.slice(0,7);
  return {
    view:          viewId,
    periodo:       label,
    fechaInicio:   fechaInicio,
    fechaFin:      fechaFin,
    todos:         readMensual(ss, 'Mensual_Todos',         fechaInicio, fechaFin, sucursal),
    local:         readMensual(ss, 'Mensual_Local',         fechaInicio, fechaFin, sucursal),
    internacional: readMensual(ss, 'Mensual_Internacional', fechaInicio, fechaFin, sucursal),
    servicios:     readServicios(ss),
    funnel:        readFunnel(ss),
    alertas:       readAlertas(ss),
    donut:         readDonut(ss),
    cashflow:      readCashFlow(ss, fechaInicio, fechaFin, sucursal),
    costos:        readCostos(ss),
    paisesOrigen:  readPaisesOrigen(ss, fechaInicio, fechaFin, sucursal),
    updated:       new Date().toISOString()
  };
}

/* ── Datos de hoja de captura — devuelve TODAS las filas (sin filtro de fechas)
   Las hojas de captura (Pacientes, Productos, Ent. Med, Estimulacion, etc.)
   muestran su contenido completo; el filtro de fechas aplica solo a datos financieros.
   ──────────────────────────────────────────────────────────────── */
function findSheet(ssCap, nombreHoja) {
  // 1. Nombre exacto
  var h = ssCap.getSheetByName(nombreHoja);
  if (h) return h;
  // 2. Búsqueda insensible a tildes y mayúsculas
  var normalize = function(s) {
    return s.trim().toLowerCase()
      .replace(/[áàäã]/g,'a').replace(/[éèë]/g,'e')
      .replace(/[íìï]/g,'i').replace(/[óòö]/g,'o')
      .replace(/[úùü]/g,'u').replace(/ñ/g,'n');
  };
  var target = normalize(nombreHoja);
  var sheets = ssCap.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalize(sheets[i].getName()) === target) return sheets[i];
  }
  return null;
}

/* Detecta la fila de encabezado de una hoja de captura (igual lógica que readCapturaData)
   y devuelve { headers, dataStart } usando posiciones absolutas de columna (incluye Periodo si existe). */
function getSheetHeaders(sheet) {
  var allRows = sheet.getDataRange().getValues();
  function countFilled(row) {
    return row.filter(function(c) { return String(c).trim() !== ''; }).length;
  }
  var r0 = allRows.length > 0 ? countFilled(allRows[0]) : 0;
  var r1 = allRows.length > 1 ? countFilled(allRows[1]) : 0;
  var headerRow, dataStart;

  if (r0 === 0) {
    headerRow = allRows[1] || [];
    dataStart = 2;
  } else if (r0 > 0 && r1 > 0) {
    var complementario = allRows[0].every(function(v, i) {
      var v0 = String(v).trim();
      var v1 = String((allRows[1][i] !== undefined ? allRows[1][i] : '')).trim();
      return !(v0 && v1);
    });
    if (complementario) {
      headerRow = allRows[0].map(function(v, i) {
        return String(v).trim() || String(allRows[1][i] !== undefined ? allRows[1][i] : '').trim();
      });
      dataStart = 2;
    } else {
      headerRow = allRows[0];
      dataStart = 1;
    }
  } else {
    headerRow = allRows[0] || [];
    dataStart = 1;
  }

  var headers = headerRow.map(function(h) { return String(h).trim(); });
  return { headers: headers, dataStart: dataStart };
}

/* insert: ?action=insert&sheet=X&Campo1=valor1&... → agrega una fila nueva al final */
function insertRow(ss, e) {
  var sheetName = (e && e.parameter.sheet) || '';
  if (SHEET_ALIASES[sheetName]) sheetName = SHEET_ALIASES[sheetName];
  if (!sheetName) return { error: 'sheet es requerido' };
  var capturaId = getCapturaId(sheetName);
  try {
    var ssIns = SpreadsheetApp.openById(capturaId);
    var shIns = findSheet(ssIns, sheetName);
    if (!shIns) return { error: 'Hoja no encontrada: ' + sheetName };
    var hdrInfo = getSheetHeaders(shIns);
    var hdrs = hdrInfo.headers;

    // Validación de duplicados: en Pacientes no se permite repetir el mismo nombre
    if (sheetName.trim().toLowerCase() === 'pacientes') {
      var nombreIdx = -1;
      for (var hi = 0; hi < hdrs.length; hi++) {
        if (hdrs[hi].toLowerCase().indexOf('nombre') > -1) { nombreIdx = hi; break; }
      }
      if (nombreIdx > -1) {
        var nombreNuevo = String(e.parameter[hdrs[nombreIdx]] || '').trim().toLowerCase();
        if (nombreNuevo) {
          var allData = shIns.getDataRange().getValues();
          for (var ri = hdrInfo.dataStart; ri < allData.length; ri++) {
            var existente = String(allData[ri][nombreIdx] || '').trim().toLowerCase();
            if (existente && existente === nombreNuevo) {
              return { error: 'Ya existe un paciente registrado con el nombre "' + e.parameter[hdrs[nombreIdx]] + '".', duplicado: true };
            }
          }
        }
      }
    }

    var newRow = hdrs.map(function(h) {
      return (h && e.parameter[h] !== undefined) ? e.parameter[h] : '';
    });
    shIns.appendRow(newRow);
    var rowNum = shIns.getLastRow();
    invalidateViewCache(sheetName);
    return { success: true, rowNum: rowNum };
  } catch(ex) {
    return { error: ex.message };
  }
}

/* ══════════════════════════════════════════════════════════════
   CAJA CHICA — hoja independiente con saldo corrido (columna TOTAL
   es fórmula en Sheets: Total_n = Total_(n-1) - Salida_n + Entrada_n).
   La pestaña activa lleva como nombre el año en curso (ej. "2026");
   muchas filas debajo de la última captura ya tienen la fórmula de
   TOTAL copiada hacia abajo en blanco, esperando captura futura —
   por eso nunca usamos appendRow() aquí, sino que buscamos la
   primera fila con FECHA y CONCEPTO vacíos para no romper el orden.
   ══════════════════════════════════════════════════════════════ */
function getCajaChicaSheet() {
  var ss = SpreadsheetApp.openById(CAJA_CHICA_SS_ID);
  var sh = ss.getSheetByName('Caja Chica');
  if (sh) return sh;
  var anioActual = String(new Date().getFullYear());
  sh = ss.getSheetByName(anioActual);
  if (sh) return sh;
  return ss.getSheets()[0]; // pestaña más reciente como último recurso
}

function readCajaChicaData() {
  try {
    var sh = getCajaChicaSheet();
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { movimientos: [], saldoInicial: 0, saldoFinal: 0 };
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iSalida   = headers.indexOf('SALIDA');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    var rows = [];
    var saldoInicial = 0;
    for (var r = 1; r < data.length; r++) {
      var row      = data[r];
      var fecha    = String(row[iFecha]    || '').trim();
      var concepto = String(row[iConcepto] || '').trim();
      if (!fecha && !concepto) continue; // fila reservada (solo fórmula de TOTAL), aún sin capturar
      var salida   = Number(row[iSalida])  || 0;
      var entrada  = Number(row[iEntrada]) || 0;
      var total    = Number(row[iTotal])   || 0;
      var esRemanente = /^REMANENTE/i.test(fecha);
      if (esRemanente) saldoInicial = total;
      rows.push({
        _rowNum:    r + 1,
        fecha:      esRemanente ? '' : fecha,
        concepto:   esRemanente ? fecha : concepto,
        esRemanente: esRemanente,
        salida:     salida,
        entrada:    entrada,
        total:      total
      });
    }

    var saldoFinal = rows.length ? rows[rows.length - 1].total : saldoInicial;

    // Resumen de gasto por periodo (admite DD/MM/YYYY y MM/DD/YYYY mezclados en la hoja)
    function parseFechaMx(f) {
      var m = f.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (!m) return null;
      var a = parseInt(m[1], 10), b = parseInt(m[2], 10), y = parseInt(m[3], 10);
      var day = a, month = b;
      if (a > 12) { day = a; month = b; }
      else if (b > 12) { day = b; month = a; }
      return new Date(y, month - 1, day);
    }
    var hoy        = new Date();
    var inicioHoy  = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    var inicio7d   = new Date(inicioHoy.getTime() - 6 * 24 * 60 * 60 * 1000);
    var inicioMes  = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    var gastoHoy = 0, gastoSemana = 0, gastoMes = 0, ingresoMes = 0;
    rows.forEach(function(m) {
      if (m.esRemanente) return;
      var d = parseFechaMx(m.fecha);
      if (!d) return;
      if (d >= inicioMes) { gastoMes += m.salida; ingresoMes += m.entrada; }
      if (d >= inicio7d)  gastoSemana += m.salida;
      if (d >= inicioHoy) gastoHoy += m.salida;
    });

    return {
      saldoInicial: saldoInicial,
      saldoFinal:   saldoFinal,
      gastoHoy:     gastoHoy,
      gastoSemana:  gastoSemana,
      gastoMes:     gastoMes,
      ingresoMes:   ingresoMes,
      movimientos:  rows.slice().reverse(), // más reciente primero
      updated:      new Date().toISOString()
    };
  } catch(ex) {
    return { error: ex.message, movimientos: [], saldoInicial: 0, saldoFinal: 0 };
  }
}

function insertCajaChicaRow(e) {
  try {
    var sh = getCajaChicaSheet();
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iSalida   = headers.indexOf('SALIDA');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    var concepto = String(e.parameter['CONCEPTO'] || '').trim();
    var fecha    = String(e.parameter['FECHA']    || '').trim();
    var salida   = parseFloat(e.parameter['SALIDA'])  || 0;
    var entrada  = parseFloat(e.parameter['ENTRADA']) || 0;

    if (!concepto)            return { error: 'El concepto es requerido.' };
    if (!fecha)               return { error: 'La fecha es requerida.' };
    if (!salida && !entrada)  return { error: 'Captura un monto de salida o entrada.' };

    // Primera fila reservada (placeholder con fórmula de TOTAL) con FECHA y CONCEPTO vacíos
    var targetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (!String(data[r][iFecha] || '').trim() && !String(data[r][iConcepto] || '').trim()) {
        targetRow = r + 1; // fila 1-indexada en la hoja
        break;
      }
    }
    if (targetRow === -1) targetRow = sh.getLastRow() + 1;

    sh.getRange(targetRow, iFecha + 1).setValue(fecha);
    sh.getRange(targetRow, iConcepto + 1).setValue(concepto);
    if (salida)  sh.getRange(targetRow, iSalida + 1).setValue(salida);
    if (entrada) sh.getRange(targetRow, iEntrada + 1).setValue(entrada);
    SpreadsheetApp.flush();

    var nuevoTotal = sh.getRange(targetRow, iTotal + 1).getValue();
    return { success: true, rowNum: targetRow, saldoFinal: Number(nuevoTotal) || 0 };
  } catch(ex) {
    return { error: ex.message };
  }
}

function getCapturaId(nombreHoja) {
  if (CAPTURA_SHEETS[nombreHoja]) return CAPTURA_SHEETS[nombreHoja];
  // Búsqueda tolerante a tildes
  var normalize = function(s) {
    return s.trim().toLowerCase()
      .replace(/[áàäã]/g,'a').replace(/[éèë]/g,'e')
      .replace(/[íìï]/g,'i').replace(/[óòö]/g,'o')
      .replace(/[úùü]/g,'u').replace(/ñ/g,'n');
  };
  var target = normalize(nombreHoja);
  var keys = Object.keys(CAPTURA_SHEETS);
  for (var i = 0; i < keys.length; i++) {
    if (normalize(keys[i]) === target) return CAPTURA_SHEETS[keys[i]];
  }
  return CAPTURA_SHEET_ID_DEFAULT;
}

function readCapturaData(ss, nombreHoja, viewId, fechaInicio, fechaFin, sucursal) {
  sucursal = sucursal || 'Todas';
  var capturaId = getCapturaId(nombreHoja);
  var ssCap = SpreadsheetApp.openById(capturaId);
  var hoja  = findSheet(ssCap, nombreHoja);
  if (!hoja) {
    return { view: viewId, headers: [], rows: [],
             error: 'Hoja "' + nombreHoja + '" no encontrada.' };
  }
  var allRows = hoja.getDataRange().getValues();
  if (allRows.length < 1) return { view: viewId, headers: [], rows: [] };

  // ── Detección inteligente de fila de encabezado ────────────────
  // Algunas hojas tienen fila 1 vacía (FET, Insumos) o encabezados
  // divididos entre fila 1 y fila 2 (ART Lab). Se detecta automáticamente.
  function countFilled(row) {
    return row.filter(function(c) { return String(c).trim() !== ''; }).length;
  }
  var r0 = countFilled(allRows[0]);
  var r1 = allRows.length > 1 ? countFilled(allRows[1]) : 0;
  var headerRow, dataStart;

  if (r0 === 0) {
    // Fila 1 vacía → encabezado en fila 2 (FET, Insumos)
    headerRow  = allRows[1] || [];
    dataStart  = 2;
  } else if (r0 > 0 && r1 > 0) {
    // Ambas filas tienen datos → verificar si son encabezados complementarios
    // (sin solapamiento de celdas llenas, patrón ART Lab)
    var complementario = allRows[0].every(function(v, i) {
      var v0 = String(v).trim();
      var v1 = String((allRows[1][i] !== undefined ? allRows[1][i] : '')).trim();
      return !(v0 && v1); // No hay posición donde ambas filas tengan valor
    });
    if (complementario) {
      // Fusionar fila 1 y fila 2 como encabezado único (ART Lab)
      headerRow = allRows[0].map(function(v, i) {
        return String(v).trim() || String(allRows[1][i] !== undefined ? allRows[1][i] : '').trim();
      });
      dataStart = 2;
    } else {
      // Fila 1 tiene encabezados reales; fila 2 es la primera fila de datos
      headerRow = allRows[0];
      dataStart = 1;
    }
  } else {
    // Caso normal: fila 1 = encabezados
    headerRow = allRows[0];
    dataStart = 1;
  }

  // Detectar columna Periodo oculta en col A (se excluye de la vista)
  var tienePeriodo = String(headerRow[0]).trim().toLowerCase() === 'periodo';
  var colStart = tienePeriodo ? 1 : 0;
  var headers = headerRow.slice(colStart)
    .map(function(h) { return String(h).trim(); })
    .filter(function(h) { return h !== ''; });

  // Incluir todas las filas no vacías (sin filtro de fechas)
  var dataRowsWithNum = allRows.slice(dataStart)
    .map(function(r, i) { return { data: r, rowNum: i + dataStart + 1 }; })
    .filter(function(item) {
      return item.data.some(function(c) { return String(c).trim() !== ''; });
    });

  var rows = dataRowsWithNum.map(function(item) {
    var r = item.data;
    var obj = { _rowNum: item.rowNum, _periodo: tienePeriodo ? String(r[0]) : '' };
    headers.forEach(function(h, i) {
      obj[h] = r[colStart + i];
      obj[h.toLowerCase()] = r[colStart + i];
    });
    return obj;
  });

  // Filtro por sucursal — solo si la columna existe y el filtro está activo
  if (sucursal && sucursal !== 'Todas') {
    var sucHdrIdx = headers.map(function(h){ return h.toLowerCase(); }).indexOf('sucursal');
    if (sucHdrIdx >= 0) {
      rows = rows.filter(function(row) {
        var val = String(row['Sucursal'] || row['sucursal'] || '').trim();
        return val === '' || val === sucursal; // vacío = hereda todas las sucursales
      });
    }
  }

  return { view: viewId, fuente: nombreHoja, headers: headers, rows: rows,
           updated: new Date().toISOString() };
}

/* ══════════════════════════════════════════════════════════════
   LECTORES INDIVIDUALES
   ══════════════════════════════════════════════════════════════ */
/* Columnas: A=Sucursal | B=Fecha(YYYY-MM-DD) | C=Mes | D=Ingresos | E=Gastos | F=Ciclos | G=CAC | H=Margen */
function readMensual(ss, sheetName, fechaInicio, fechaFin, sucursal) {
  var hoja = ss.getSheetByName(sheetName);
  if (!hoja) return { meses:[], ingresos:[], gastos:[], ciclos:[], cac:[], margen:[] };
  var rows = hoja.getDataRange().getValues();
  if (rows.length < 2) return { meses:[], ingresos:[], gastos:[], ciclos:[], cac:[], margen:[] };
  // Detectar columnas por encabezado (soporta columnas en cualquier orden)
  var hdrs = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iSuc    = hdrs.indexOf('sucursal');
  var iFecha  = hdrs.indexOf('fecha');  if (iFecha  < 0) iFecha  = 1;
  var iMes    = hdrs.indexOf('mes');    if (iMes    < 0) iMes    = 2;
  var iIngr   = hdrs.indexOf('ingresos'); if (iIngr < 0) iIngr   = 3;
  var iGast   = hdrs.indexOf('gastos');   if (iGast < 0) iGast   = 4;
  var iCiclos = hdrs.indexOf('ciclos');   if (iCiclos < 0) iCiclos = 5;
  var iCac    = hdrs.indexOf('cac');      if (iCac  < 0) iCac    = 6;
  var iMargen = hdrs.indexOf('margen');   if (iMargen < 0) iMargen = 7;
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[iFecha]).trim();
    if (f < fechaInicio || f > fechaFin) return false;
    if (sucursal && sucursal !== 'Todas' && iSuc >= 0) {
      var s = String(r[iSuc] || '').trim();
      if (s && s !== sucursal) return false;
    }
    return true;
  });
  data.sort(function(a, b) { return String(a[iFecha]) < String(b[iFecha]) ? -1 : 1; });
  return {
    meses:    data.map(function(r) { return String(r[iMes]); }),
    ingresos: data.map(function(r) { return Number(r[iIngr]); }),
    gastos:   data.map(function(r) { return Number(r[iGast]); }),
    ciclos:   data.map(function(r) { return Number(r[iCiclos]); }),
    cac:      data.map(function(r) { return Number(r[iCac]); }),
    margen:   data.map(function(r) { return Number(r[iMargen]); })
  };
}

function readSucursales(ss) {
  var hoja = ss.getSheetByName('Sucursales');
  if (!hoja) return [];
  var rows = hoja.getDataRange().getValues();
  if (rows.length < 2) return [];
  var hdrs = rows[0].map(function(h){ return String(h).trim(); });
  return rows.slice(1)
    .filter(function(r){ return String(r[0]).trim() !== ''; })
    .map(function(r){
      var obj = {};
      hdrs.forEach(function(h, i){ obj[h] = r[i]; });
      return obj;
    })
    .filter(function(s){
      var activo = String(s['Activo'] !== undefined ? s['Activo'] : 'true').trim().toLowerCase();
      return activo !== 'false' && activo !== 'no' && activo !== '0';
    });
}

/* ── Funciones financieras definidas en api_finance.gs (mismo proyecto GAS) ──
   readCashFlow, readServicios, readFunnel, readAlertas, readDonut,
   readPaisesOrigen, readCostos, readEstadoResultados, readOperatingPL,
   _buildPLReport, readBanksData, saveBankRow, doPost
   ────────────────────────────────────────────────────────────────────────── */