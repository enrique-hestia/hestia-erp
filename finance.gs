/* ==============================================================
   finance.gs — Módulo Financiero
   --------------------------------------------------------------
   CashFlow, P&L, ER, Bancos (3 cuentas), CxC/CxP futuros
   Proyecto Google Apps Script — Hestia Fertility ERP
   Todas las constantes vienen de config.gs (mismo proyecto)
   ============================================================== */

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

  // Helper: busca un valor en erLookup probando múltiples variantes del label
  function _look(keys) {
    for (var ki=0;ki<keys.length;ki++) {
      var v=erLookup[normalize(keys[ki])];
      if (v && (v.actual!==0||v.prev!==0)) return v;
    }
    return {actual:0,prev:0};
  }
  function _lookB(keys) {
    for (var ki=0;ki<keys.length;ki++) {
      var v=bgLookup[normalize(keys[ki])];
      if (v && v!==0) return v;
    }
    return 0;
  }
  function _sum(labels) {
    var a=0,p=0;
    labels.forEach(function(lbl){var v=erLookup[normalize(lbl)]||{actual:0,prev:0};a+=Math.abs(v.actual);p+=Math.abs(v.prev);});
    return {actual:a,prev:p};
  }
  function _sumB(labels) {
    var t=0;
    labels.forEach(function(lbl){t+=Math.abs(bgLookup[normalize(lbl)]||0);});
    return t;
  }

  // Revenue y COGS — desde ER sheet
  var _rev  = _look(['total income','total revenue','ingresos totales','total ingresos']);
  var _cogs = _look(['total cogs','total cost of goods','costo de ventas','total costos']);

  // OpEx — intenta lookup, si no suma hijos
  var _opexLook = _look(['total opex','total operating expenses','gastos operativos','total gastos operativos','opex']);
  if (_opexLook.actual===0) _opexLook = _sum(['marketing','renta','mto renta','servicios','mto. renta','mantenimiento renta']);
  var _opex = _opexLook;

  // G&A — intenta lookup, si no suma hijos
  var _gaLook = _look(['total g&a','total g & a','g&a','general & administrative','general and administrative','administracion','administración','gastos administracion']);
  if (_gaLook.actual===0) _gaLook = _sum(['nomina','nómina','gastos varios','gastos generales']);
  var _ga = _gaLook;

  var _da  = _look(['depreciation & amortization','depreciacion','depreciación','d&a','da']);
  var _tax = _look(['taxes','impuestos','isr','isr e ietu']);

  // Métricas derivadas en cascada
  function _deriv(a, p) { return {actual:a, prev:p}; }
  var _gp     = _deriv(_rev.actual - Math.abs(_cogs.actual), _rev.prev - Math.abs(_cogs.prev));
  var _cc     = _deriv(_gp.actual  - Math.abs(_opex.actual), _gp.prev  - Math.abs(_opex.prev));
  var _ebitda = _deriv(_cc.actual  - Math.abs(_ga.actual),   _cc.prev  - Math.abs(_ga.prev));
  var _ebit   = _deriv(_ebitda.actual - Math.abs(_da.actual), _ebitda.prev - Math.abs(_da.prev));
  var _ebt    = _ebit;
  var _np     = _deriv(_ebt.actual - Math.abs(_tax.actual),  _ebt.prev - Math.abs(_tax.prev));

  // Inyectar en lookup (solo si falta o es 0)
  function _inject(k, val) { if (!erLookup[k] || erLookup[k].actual===0) erLookup[k]=val; }
  _inject('total income',        _rev);
  _inject('total cogs',          _cogs);
  _inject('total opex',          _opex);
  _inject('total g&a',           _ga);
  _inject('gross profit',        _gp);
  _inject('clinic contribution', _cc);
  _inject('ebitda',              _ebitda);
  _inject('ebit',                _ebit);
  _inject('ebt',                 _ebt);
  _inject('net profit',          _np);

  // Budget — misma lógica
  var _bRev  = _lookB(['total income','total revenue','ingresos totales']);
  var _bCogs = _lookB(['total cogs','total cost of goods','costo de ventas']);
  var _bOpex = _lookB(['total opex','total operating expenses','gastos operativos','opex']);
  if (!_bOpex) _bOpex = _sumB(['marketing','renta','mto renta','servicios']);
  var _bGa   = _lookB(['total g&a','total g & a','g&a','administracion','administración']);
  if (!_bGa)   _bGa   = _sumB(['nomina','nómina','gastos varios']);
  var _bDa   = _lookB(['depreciation & amortization','depreciacion','depreciación']);
  var _bTax  = _lookB(['taxes','impuestos','isr']);
  var _bGp     = _bRev  - Math.abs(_bCogs);
  var _bCc     = _bGp   - Math.abs(_bOpex);
  var _bEbitda = _bCc   - Math.abs(_bGa);
  var _bEbit   = _bEbitda - Math.abs(_bDa);
  var _bNp     = _bEbit   - Math.abs(_bTax);
  function _injectB(k, val) { if (!bgLookup[k]||bgLookup[k]===0) bgLookup[k]=val; }
  _injectB('total income',        _bRev);
  _injectB('total cogs',          _bCogs);
  _injectB('total opex',          _bOpex);
  _injectB('total g&a',           _bGa);
  _injectB('gross profit',        _bGp);
  _injectB('clinic contribution', _bCc);
  _injectB('ebitda',              _bEbitda);
  _injectB('ebit',                _bEbit);
  _injectB('ebt',                 _bEbit);
  _injectB('net profit',          _bNp);

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
            }),
            erKeys:Object.keys(erLookup).slice(0,40)}
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
    if (body.action === 'saveBankRow') {
      var result = saveBankRow(body.banco, body.row);
      try { CacheService.getScriptCache().remove('erp_banks_v1'); } catch(e) {}
      return jsonResponse(result);
    }
    if (body.action === 'deleteBankRow') {
      var result = deleteBankRow(body.banco, body.rowNum);
      try { CacheService.getScriptCache().remove('erp_banks_v1'); } catch(e) {}
      return jsonResponse(result);
    }
    if (body.action === 'updateBankRow') {
      var result = updateBankRow(body.banco, body.rowNum, body.row);
      try { CacheService.getScriptCache().remove('erp_banks_v1'); } catch(e) {}
      return jsonResponse(result);
    }
    if (body.action === 'createBankSheet') {
      var result = createBankSheet(body.nombre, body.color);
      return jsonResponse(result);
    }
    if (body.action === 'saveLiberado') {
      var result = saveLiberado(body.rowNum, body.liberado);
      try { CacheService.getScriptCache().remove('erp_banks_v1'); } catch(e) {}
      return jsonResponse(result);
    }
    if (body.action === 'saveIngreso') {
      return jsonResponse(saveIngreso(body));
    }
    if (body.action === 'uploadIngresoPDF') {
      return jsonResponse(uploadIngresoPDF(body.opId, body.tipo||'factura', body.fileName, body.base64, body.mimeType));
    }
    if (body.action === 'setupBDIngresos') {
      return jsonResponse(setupBDIngresos());
    }
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
      B.movimientos=r.slice(1).map(function(x,idx){
        var d=num(x[1]),t=num(x[2]);
        return{rowNum:idx+2,fecha:dt(x[0]),deposito:d,retiro:t,monto:d>0?d:-t,saldo:num(x[3]),
               referencia:String(x[4]||''),depositoUSD:num(x[5]),tipoCambio:num(x[6]),
               poliza:String(x[7]||''),observaciones:String(x[8]||''),tipo:d>0?'deposito':'retiro'};
      }).slice(-30).reverse();
      return B;
    }
    function rAmex(sheet) {
      var B={id:'amex',nombre:'AMEX',color:'#007bc1',saldo:0,movimientos:[],totalRows:0};
      if(!sheet) return B;
      var r=sheet.getDataRange().getValues(); if(r.length<2) return B;
      for(var i=r.length-1;i>=1;i--){ var s=num(r[i][2]); if(s!==0){B.saldo=s;break;} }
      B.totalRows=r.length-1;
      B.movimientos=r.slice(1).map(function(x,idx){
        var m=num(x[1]);
        return{rowNum:idx+2,fecha:dt(x[0]),monto:m,saldo:num(x[2]),referencia:String(x[3]||''),
               usd:num(x[4]),tipoCambio:num(x[5]),notas:String(x[6]||''),
               poliza:String(x[7]||''),mes:String(x[8]||''),tipo:m>=0?'cargo':'pago'};
      }).slice(-30).reverse();
      return B;
    }
    function rMP(sheet) {
      var B={id:'mercadopago',nombre:'Mercado Pago',color:'#009ee3',saldo:0,porLiberar:0,movimientos:[],totalRows:0};
      if(!sheet) return B;
      var r=sheet.getDataRange().getValues(); if(r.length<2) return B;
      var saldoLib=0, porLib=0;
      for(var i=1;i<r.length;i++){
        var tv=num(r[i][5]);
        var lib=r[i][7]===true||String(r[i][7]).toUpperCase()==='TRUE';
        if(lib) saldoLib+=tv;
        else if(tv>0) porLib+=tv;
      }
      B.saldo=saldoLib;
      B.porLiberar=porLib;
      // Filtrar solo filas con fecha válida o cobro real (excluye filas de fórmula vacías)
      var dataRows=[];
      for(var i=1;i<r.length;i++){
        var hasFecha=r[i][1]&&r[i][1]!=='';
        var hasCobro=num(r[i][2])!==0||num(r[i][5])!==0;
        if(hasFecha||hasCobro) dataRows.push({row:r[i],sheetRow:i+1});
      }
      B.totalRows=dataRows.length;
      var _mNames=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      function fmtMes(v){
        if(v instanceof Date) return _mNames[v.getMonth()]+'-'+v.getFullYear();
        var s=String(v||'').trim();
        var d=new Date(s); if(!isNaN(d)&&s.length>6) return _mNames[d.getMonth()]+'-'+d.getFullYear();
        return s;
      }
      B.movimientos=dataRows.slice(-30).reverse().map(function(entry){
        var x=entry.row;
        return{rowNum:entry.sheetRow,mes:fmtMes(x[0]),fecha:dt(x[1]),cobro:num(x[2]),comisiones:num(x[3]),
               pctComision:num(x[4]),totalVenta:num(x[5]),saldo:num(x[6]),
               liberado:x[7]===true||String(x[7]).toUpperCase()==='TRUE',
               observaciones:String(x[8]||''),tipo:String(x[9]||'CARGO').toUpperCase(),monto:num(x[5])};
      });
      // Resumen de comisiones por mes (todas las filas, no solo últimas 30)
      var comByMes = {};
      for(var ci=0;ci<dataRows.length;ci++){
        var cx=dataRows[ci].row;
        var mes=fmtMes(cx[0]); if(!mes) continue;
        var cobro=num(cx[2]), com=Math.abs(num(cx[3])), neto=num(cx[5]);
        if(!comByMes[mes]) comByMes[mes]={mes:mes,totalCobro:0,totalComision:0,totalNeto:0,movimientos:0};
        if(cobro>0) comByMes[mes].totalCobro+=cobro;
        comByMes[mes].totalComision+=com;
        if(neto>0) comByMes[mes].totalNeto+=neto;
        comByMes[mes].movimientos++;
      }
      B.comisionesPorMes=[];
      for(var mk in comByMes){
        var m=comByMes[mk];
        m.pctPromedio=m.totalCobro>0?(m.totalComision/m.totalCobro):0;
        B.comisionesPorMes.push(m);
      }
      // Ordenar: mes actual primero, más antiguo al final
      var _mIdx={Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
      function mesNum(s){var p=String(s||'').split('-');var mm=parseInt(p[0])||_mIdx[p[0]]||0;var yy=parseInt(p[1])||0;return yy*100+mm;}
      B.comisionesPorMes.sort(function(a,b){return mesNum(b.mes)-mesNum(a.mes);});
      return B;
    }

    // Bancos estándar
    var sant=rSant(byGid(BANKS_GID.santander));
    var amex=rAmex(byGid(BANKS_GID.amex));
    var mp  =rMP(byGid(BANKS_GID.mercadopago));
    var bancos={santander:sant,amex:amex,mercadopago:mp};

    // Bancos adicionales (cualquier clave en BANKS_GID que no sea los tres estándar)
    var std=['santander','amex','mercadopago'];
    for (var key in BANKS_GID) {
      if (std.indexOf(key) >= 0) continue;
      var extraSheet = byGid(BANKS_GID[key]);
      if (!extraSheet) continue;
      // Lector genérico: Fecha|Depósito|Retiro|Saldo|Referencia|...
      var rExtra = extraSheet.getDataRange().getValues();
      var EB = {id:key, nombre:extraSheet.getName(), color:'#6b7280', saldo:0, movimientos:[], totalRows:0};
      if (rExtra.length > 1) {
        for (var ei=rExtra.length-1;ei>=1;ei--){ var es=num(rExtra[ei][3]); if(es!==0){EB.saldo=es;break;} }
        EB.totalRows = rExtra.length-1;
        EB.movimientos = rExtra.slice(1).map(function(x,idx){
          var d=num(x[1]),t=num(x[2]);
          return{rowNum:idx+2,fecha:dt(x[0]),deposito:d,retiro:t,monto:d>0?d:-t,
                 saldo:num(x[3]),referencia:String(x[4]||''),tipo:d>0?'deposito':'retiro'};
        }).slice(-30).reverse();
      }
      bancos[key]=EB;
    }

    var totalSaldo=0;
    for (var bk in bancos) totalSaldo+=bancos[bk].saldo||0;
    return { view:'cashflow', fuente:'CashFlow', bancos:bancos, totalSaldo:totalSaldo };
  } catch(ex) {
    return { view:'cashflow', fuente:'CashFlow', error:ex.message, bancos:{}, totalSaldo:0 };
  }
}

function deleteBankRow(banco, rowNum) {
  try {
    var ss=SpreadsheetApp.openById(BANKS_SS_ID);
    var sh=ss.getSheets();
    var key=String(banco).toLowerCase().replace(/[\s-]/g,'');
    var gid=BANKS_GID[key];
    if (gid===undefined) return {ok:false, error:'Banco desconocido: '+banco};
    var sheet=null;
    for(var i=0;i<sh.length;i++) if(sh[i].getSheetId()===gid){sheet=sh[i];break;}
    if(!sheet) return {ok:false, error:'Pestaña no encontrada'};
    if(rowNum<2||rowNum>sheet.getLastRow()) return {ok:false, error:'Fila fuera de rango'};
    sheet.deleteRow(rowNum);
    // Recalcular saldos acumulados después del borrado
    _recalcSaldos(sheet, key);
    return {ok:true, banco:banco, deletedRow:rowNum};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function updateBankRow(banco, rowNum, row) {
  try {
    var ss=SpreadsheetApp.openById(BANKS_SS_ID);
    var sh=ss.getSheets();
    var key=String(banco).toLowerCase().replace(/[\s-]/g,'');
    var gid=BANKS_GID[key];
    if (gid===undefined) return {ok:false, error:'Banco desconocido: '+banco};
    var sheet=null;
    for(var i=0;i<sh.length;i++) if(sh[i].getSheetId()===gid){sheet=sh[i];break;}
    if(!sheet) return {ok:false, error:'Pestaña no encontrada'};
    if(rowNum<2||rowNum>sheet.getLastRow()) return {ok:false, error:'Fila fuera de rango'};
    var range=sheet.getRange(rowNum,1,1,row.length);
    range.setValues([row]);
    _recalcSaldos(sheet, key);
    return {ok:true, banco:banco, updatedRow:rowNum};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

// Recalcula la columna de saldo acumulado tras editar/borrar
function _recalcSaldos(sheet, key) {
  var lr=sheet.getLastRow(); if(lr<2) return;
  function num(v){var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}
  if(key==='santander') {
    var vals=sheet.getRange(2,1,lr-1,4).getValues();
    var run=0;
    for(var i=0;i<vals.length;i++){run+=num(vals[i][1])-num(vals[i][2]); vals[i][3]=run;}
    sheet.getRange(2,4,lr-1,1).setValues(vals.map(function(r){return[r[3]];}));
  } else if(key==='amex') {
    var vals=sheet.getRange(2,1,lr-1,3).getValues();
    var run=0;
    for(var i=0;i<vals.length;i++){run+=num(vals[i][1]); vals[i][2]=run;}
    sheet.getRange(2,3,lr-1,1).setValues(vals.map(function(r){return[r[2]];}));
  } else if(key==='mercadopago') {
    // Recalcula col G (saldo corrido) = suma acumulada de col F (Total de Venta)
    var vals=sheet.getRange(2,6,lr-1,1).getValues(); // col F = totalVenta (index 5, col 6)
    var run=0;
    var saldos=vals.map(function(r){run+=num(r[0]); return[run];});
    sheet.getRange(2,7,lr-1,1).setValues(saldos); // col G = saldo
  }
}

function createBankSheet(nombre, color) {
  try {
    var ss=SpreadsheetApp.openById(BANKS_SS_ID);
    // Verificar que no exista ya
    var sheets=ss.getSheets();
    var key=nombre.toLowerCase().replace(/[\s-]/g,'');
    for(var i=0;i<sheets.length;i++) if(sheets[i].getName().toLowerCase().replace(/[\s-]/g,'')=== key) return {ok:false,error:'Ya existe una pestaña con ese nombre'};
    var newSheet=ss.insertSheet(nombre);
    // Headers: Fecha | Depósito | Retiro | Saldo | Referencia | USD | T.Cambio | Póliza | Observaciones
    newSheet.getRange(1,1,1,9).setValues([['Fecha','Depósito','Retiro','Saldo','Referencia','USD','T.Cambio','Póliza','Observaciones']]);
    newSheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#f3f4f6');
    newSheet.setFrozenRows(1);
    var gid=newSheet.getSheetId();
    return {ok:true, nombre:nombre, gid:gid, key:key,
            instruccion:'Agrega esta línea a BANKS_GID en config.gs: '+key+': '+gid};
  } catch(ex) { return {ok:false, error:ex.message}; }
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
    if (key==='santander') {
      row[3]=ls+(parseFloat(row[1])||0)-(parseFloat(row[2])||0);
    } else if (key==='amex') {
      row[2]=ls+(parseFloat(row[1])||0);
    } else {
      // MP col E (idx 4): pct = 1-(totalVenta/cobro) como fracción decimal (ej. 0.0406)
      var mpCobro = parseFloat(row[2]) || 0;
      var mpNeto  = parseFloat(row[5]) || 0;
      row[4] = (mpCobro !== 0) ? (1 - (mpNeto / mpCobro)) : 0;
      // MP col G (idx 6): saldo corrido = suma acumulada de col F
      var allVals=sheet.getLastRow()>1?sheet.getRange(2,6,sheet.getLastRow()-1,1).getValues():[];
      var runSum=0; for(var k=0;k<allVals.length;k++) runSum+=parseFloat(allVals[k][0])||0;
      row[6]=runSum+(parseFloat(row[5])||0);
    }
    sheet.appendRow(row);
    return {ok:true, banco:banco, newSaldo:row[sc-1], totalRows:sheet.getLastRow()-1};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function saveLiberado(rowNum, liberado) {
  try {
    var ss=SpreadsheetApp.openById(BANKS_SS_ID);
    var sh=ss.getSheets();
    var sheet=null;
    for(var i=0;i<sh.length;i++) if(sh[i].getSheetId()===BANKS_GID.mercadopago){sheet=sh[i];break;}
    if(!sheet) return {ok:false,error:'Hoja Mercado Pago no encontrada'};
    if(rowNum<2||rowNum>sheet.getLastRow()) return {ok:false,error:'Fila fuera de rango'};
    sheet.getRange(rowNum,8).setValue(liberado===true||liberado==='true'||liberado===1);
    return {ok:true,rowNum:rowNum,liberado:liberado};
  } catch(ex) { return {ok:false,error:ex.message}; }
}

/* ══════════════════════════════════════════════════════════════
   CUENTAS POR PAGAR — lectura de hoja externa, solo reporte
   ══════════════════════════════════════════════════════════════ */
function readCxPData() {
  try {
    var ss    = SpreadsheetApp.openById(CXP_SS_ID);
    var sheets= ss.getSheets();
    var sheet = null;
    for (var i=0;i<sheets.length;i++) if(sheets[i].getSheetId()===CXP_GID){sheet=sheets[i];break;}
    if (!sheet) sheet = ss.getSheets()[0];
    var raw = sheet.getDataRange().getValues();
    if (raw.length < 2) return {view:'cxp',headers:[],rows:[],resumen:{vencido:0,d3:0,d7:0,d30:0,totalVencido:0,totalD3:0,totalD7:0,totalD30:0}};

    var headers = raw[0].map(function(h){ return String(h).trim(); });
    // Detectar columna de vencimiento: busca por nombre o cae en columna L (índice 11)
    var iVenc = -1;
    var vencKw = /venc|due|expir|plazo/i;
    for (var c=0;c<headers.length;c++) { if(vencKw.test(headers[c])){iVenc=c;break;} }
    if (iVenc < 0) iVenc = Math.min(11, headers.length-1); // columna L

    // Columna de monto: siempre columna J (índice 9)
    var iMonto = Math.min(9, headers.length-1);

    var hoy = new Date();
    hoy.setHours(0,0,0,0);

    function parseDate(v) {
      if (!v) return null;
      if (v instanceof Date) { var d=new Date(v); d.setHours(0,0,0,0); return d; }
      var s=String(v).trim();
      // DD/MM/YYYY
      var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if(m){ return new Date(parseInt(m[3]),parseInt(m[2])-1,parseInt(m[1])); }
      // YYYY-MM-DD
      m=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if(m){ return new Date(parseInt(m[1]),parseInt(m[2])-1,parseInt(m[3])); }
      var d=new Date(s); return isNaN(d)?null:d;
    }

    function diasRestantes(fechaVenc) {
      if (!fechaVenc) return null;
      return Math.round((fechaVenc - hoy) / 86400000);
    }

    function urgencia(dias) {
      if (dias === null) return 'sin-fecha';
      if (dias < 0)  return 'vencido';
      if (dias <= 3) return 'd3';
      if (dias <= 7) return 'd7';
      if (dias <= 30) return 'd30';
      return 'ok';
    }

    var resumen = {vencido:0,d3:0,d7:0,d30:0,totalVencido:0,totalD3:0,totalD7:0,totalD30:0};
    var rows = [];

    for (var r=1;r<raw.length;r++) {
      var row = raw[r];
      // Saltar filas completamente vacías
      if (row.every(function(c){ return String(c).trim()===''; })) continue;
      var fechaVenc = parseDate(row[iVenc]);
      var dias = diasRestantes(fechaVenc);
      var urg  = urgencia(dias);
      var monto = iMonto>=0 ? (parseFloat(row[iMonto])||0) : 0;

      if (urg==='vencido') { resumen.vencido++; resumen.totalVencido+=monto; }
      else if (urg==='d3') { resumen.d3++;      resumen.totalD3+=monto; }
      else if (urg==='d7') { resumen.d7++;      resumen.totalD7+=monto; }
      else if (urg==='d30'){ resumen.d30++;     resumen.totalD30+=monto; }

      var obj = {_urgencia:urg, _dias:dias, _monto:monto};
      headers.forEach(function(h,i){
        var v=row[i];
        if (v instanceof Date) {
          var dd=String(v.getDate()).padStart(2,'0');
          var mm=String(v.getMonth()+1).padStart(2,'0');
          obj[h] = dd+'/'+mm+'/'+v.getFullYear();
        } else {
          obj[h] = (v===null||v===undefined)?'':v;
        }
      });
      rows.push(obj);
    }

    // Ordenar: vencidos primero, luego por días ascendente
    rows.sort(function(a,b){
      var ua={'vencido':0,'d3':1,'d7':2,'d30':3,'ok':4,'sin-fecha':5}[a._urgencia]||4;
      var ub={'vencido':0,'d3':1,'d7':2,'d30':3,'ok':4,'sin-fecha':5}[b._urgencia]||4;
      if(ua!==ub) return ua-ub;
      return (a._dias||999)-(b._dias||999);
    });

    return {view:'cxp', headers:headers, rows:rows, resumen:resumen,
            iVenc:iVenc, iMonto:iMonto, updated:new Date().toISOString()};
  } catch(ex) {
    return {view:'cxp', error:ex.message, headers:[], rows:[], resumen:{}};
  }
}

/* ══════════════════════════════════════════════════════════════
   INGRESOS — lectura de spreadsheet mensual de ingresos
   ══════════════════════════════════════════════════════════════ */
var INGRESOS_SS_ID = '1x_TE_YxLOwnBXKV_lA3Ss_EOSu1p61uTdUmw2zh_6uc'; // Ingresos 2026
var INGRESOS_SS_2025 = '17gNzXavMbQ8DhFEIxCqzCJ6z-wTZgIwL4ibEDyVKE2w';
var INGRESOS_SS_2024 = '1Zx4QWulAgrrVBeI8nfTR10EiYL-l3crJsaZSYDbWSug';

function _readFromBDIngresos(sheet) {
  var raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return {view:'ingresos',rows:[],resumenMensual:[],totalAnual:0,totalPagado:0,numOperaciones:0,ticketPromedio:0,topCategorias:[],topFormasPago:[],totalRows:0,headers:[]};

  function num(v){if(typeof v==='number')return v;var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}
  function dt(v){if(!v)return'';if(v instanceof Date)return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');return String(v);}

  var MESES_MAP = {'01':'Enero','02':'Febrero','03':'Marzo','04':'Abril','05':'Mayo','06':'Junio',
                   '07':'Julio','08':'Agosto','09':'Septiembre','10':'Octubre','11':'Noviembre','12':'Diciembre'};
  var MESES_IDX = {Enero:0,Febrero:1,Marzo:2,Abril:3,Mayo:4,Junio:5,Julio:6,Agosto:7,Septiembre:8,Octubre:9,Noviembre:10,Diciembre:11};

  var allRows = [];
  var opTotals = {}; // OP -> sum of TotalLinea

  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var op = String(r[0]||'').trim();
    if (!op) continue;

    var fecha = dt(r[2]);
    var mesNum = fecha.substring(5,7);
    var mesName = MESES_MAP[mesNum] || '';
    var mesIdx = MESES_IDX[mesName]; if (mesIdx === undefined) mesIdx = -1;
    var totalLinea = num(r[9]);

    if (!opTotals[op]) opTotals[op] = 0;
    opTotals[op] += totalLinea;

    // BD_Ingresos cols: OP(0),Linea(1),Fecha(2),Paciente(3),Cat(4),Prod(5),
    // PVP(6),Desc(7),Cant(8),TotalPagar(9),Pagado(10),MontoFactMes(11),
    // FormaPago(12),Facturacion(13),Conciliacion(14),Contabilidad(15),
    // Obs(16),Factura(17),Poliza(18),USMX(19),Ciclo(20),Sucursal(21),ArchivoURL(22)
    allRows.push({
      id: op,
      linea: num(r[1]),
      fecha: fecha,
      paciente: String(r[3]||''),
      categoria: String(r[4]||''),
      producto: String(r[5]||''),
      pvp: num(r[6]),
      descuento: num(r[7]),
      cantidad: num(r[8]),
      totalPagar: totalLinea,
      pagado: num(r[10]),
      montoFact: num(r[11]),
      formaPago: String(r[12]||''),
      facturacion: r[13]===true||String(r[13]).toUpperCase()==='TRUE',
      conciliacion: r[14]===true||String(r[14]).toUpperCase()==='TRUE',
      contabilidad: r[15]===true||String(r[15]).toUpperCase()==='TRUE',
      observaciones: String(r[16]||''),
      factura: String(r[17]||''),
      poliza: String(r[18]||''),
      moneda: String(r[19]||''),
      ciclo: String(r[20]||''),
      sucursal: String(r[21]||''),
      archivoURL: String(r[22]||''),
      mes: mesName,
      mesIdx: mesIdx
    });
  }

  // KPIs y resúmenes
  var totalAnual=0,totalPagado=0,cP=0,sP=0;
  var catMap={},fpMap={},mesMap={};
  for (var ai=0;ai<allRows.length;ai++){
    var ar=allRows[ai];
    totalAnual+=ar.totalPagar; totalPagado+=ar.pagado;
    if(ar.totalPagar>0){cP++;sP+=ar.totalPagar;}
    var c=ar.categoria||'Sin categoría';
    if(!catMap[c])catMap[c]={nombre:c,total:0,count:0};catMap[c].total+=ar.totalPagar;catMap[c].count++;
    var f=ar.formaPago||'Sin especificar';
    if(!fpMap[f])fpMap[f]={nombre:f,total:0,count:0};fpMap[f].total+=ar.totalPagar;fpMap[f].count++;
    var mn=ar.mes;
    if(mn&&!mesMap[mn])mesMap[mn]={mes:mn,mesIdx:ar.mesIdx,totalIngresos:0,totalPagado:0,numOperaciones:0};
    if(mn){mesMap[mn].totalIngresos+=ar.totalPagar;mesMap[mn].totalPagado+=ar.pagado;mesMap[mn].numOperaciones++;}
  }

  function sortTop(map,limit){
    var arr=[];for(var k in map)arr.push(map[k]);
    arr.sort(function(a,b){return b.total-a.total;});
    for(var i=0;i<arr.length;i++)arr[i].pct=totalAnual>0?(arr[i].total/totalAnual)*100:0;
    return arr.slice(0,limit);
  }
  var rmArr=[];for(var mk in mesMap)rmArr.push(mesMap[mk]);
  rmArr.sort(function(a,b){return a.mesIdx-b.mesIdx;});

  // Operaciones únicas para contar
  var uniqueOps={};for(var ui=0;ui<allRows.length;ui++)uniqueOps[allRows[ui].id]=true;
  var numOps=Object.keys(uniqueOps).length;

  return {
    view:'ingresos', fuente:'BD_Ingresos',
    anio: allRows.length>0 ? allRows[0].fecha.substring(0,4) : String(new Date().getFullYear()),
    rows: allRows.slice(-500).reverse(),
    totalRows: allRows.length,
    resumenMensual: rmArr,
    totalAnual: totalAnual,
    totalPagado: totalPagado,
    numOperaciones: numOps,
    ticketPromedio: numOps>0?totalAnual/numOps:0,
    topCategorias: sortTop(catMap,10),
    topFormasPago: sortTop(fpMap,10),
    headers:['OP','Fecha','Paciente','Categoría','Producto','P.V.P.','Cant.','Total','F.Pago','Sucursal']
  };
}

function readIngresosData() {
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);

    // Intentar leer de BD_Ingresos primero
    var bdSheet = null;
    var allSheets = ss.getSheets();
    for (var bi = 0; bi < allSheets.length; bi++) {
      if (allSheets[bi].getName() === BD_INGRESOS_TAB) { bdSheet = allSheets[bi]; break; }
    }
    if (bdSheet && bdSheet.getLastRow() > 1) {
      return _readFromBDIngresos(bdSheet);
    }
    // Fallback: leer de pestañas mensuales
    var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                 'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    var tz = Session.getScriptTimeZone();

    function num(v) {
      if (typeof v === 'number') return v;
      var n = parseFloat(String(v || '').replace(/[$,\s]/g, ''));
      return isNaN(n) ? 0 : n;
    }
    function dt(v) {
      if (!v) return '';
      if (v instanceof Date) {
        return v.getFullYear() + '-' +
          String(v.getMonth() + 1).padStart(2, '0') + '-' +
          String(v.getDate()).padStart(2, '0');
      }
      return String(v);
    }

    var allRows = [];
    var monthSummaries = [];

    var sheets = ss.getSheets();
    for (var si = 0; si < sheets.length; si++) {
      var tabName = sheets[si].getName().trim();
      var mesIdx = MESES.indexOf(tabName);
      if (mesIdx < 0) continue;

      var raw = sheets[si].getDataRange().getValues();
      if (raw.length < 2) continue;

      var tabRows = [];
      for (var ri = 1; ri < raw.length; ri++) {
        var r = raw[ri];
        var idVal = r[0];
        // Skip empty or non-numeric ID rows (summaries/formulas)
        if (idVal === '' || idVal === null || idVal === undefined) continue;
        var idNum = Number(idVal);
        if (isNaN(idNum) || idNum <= 0) continue;

        var totalPagar = num(r[8]);
        var pagado = num(r[9]);

        tabRows.push({
          id: idVal,
          fecha: dt(r[1]),
          paciente: String(r[2] || ''),
          categoria: String(r[3] || ''),
          producto: String(r[4] || ''),
          pvp: num(r[5]),
          descuento: num(r[6]),
          cantidad: num(r[7]),
          totalPagar: totalPagar,
          pagado: pagado,
          montoFact: num(r[10]),
          formaPago: String(r[11] || ''),
          facturacion: String(r[12] || ''),
          conciliacion: String(r[13] || ''),
          contabilidad: String(r[14] || ''),
          observaciones: String(r[15] || ''),
          factura: String(r[16] || ''),
          poliza: String(r[17] || ''),
          moneda: String(r[18] || ''),
          ciclo: String(r[19] || ''),
          sucursal: (r.length > 20) ? String(r[20] || '') : '',
          mes: tabName,
          mesIdx: mesIdx
        });
      }

      // Month summary
      var totalIngresos = 0, totalPagadoMes = 0, countPositive = 0, sumPositive = 0;
      var categorias = {}, formasPago = {};
      for (var ti = 0; ti < tabRows.length; ti++) {
        var row = tabRows[ti];
        totalIngresos += row.totalPagar;
        totalPagadoMes += row.pagado;
        if (row.totalPagar > 0) { countPositive++; sumPositive += row.totalPagar; }

        var cat = row.categoria || 'Sin categoría';
        if (!categorias[cat]) categorias[cat] = { total: 0, count: 0 };
        categorias[cat].total += row.totalPagar;
        categorias[cat].count++;

        var fp = row.formaPago || 'Sin especificar';
        if (!formasPago[fp]) formasPago[fp] = { total: 0, count: 0 };
        formasPago[fp].total += row.totalPagar;
        formasPago[fp].count++;
      }

      monthSummaries.push({
        mes: tabName,
        mesIdx: mesIdx,
        totalIngresos: totalIngresos,
        totalPagado: totalPagadoMes,
        numOperaciones: tabRows.length,
        ticketPromedio: countPositive > 0 ? sumPositive / countPositive : 0,
        categorias: categorias,
        formasPago: formasPago
      });

      allRows = allRows.concat(tabRows);
    }

    // Sort month summaries by mesIdx
    monthSummaries.sort(function(a, b) { return a.mesIdx - b.mesIdx; });

    // Top-level KPIs
    var totalAnual = 0, totalPagado = 0, numOps = allRows.length;
    var countPos = 0, sumPos = 0;
    var catMap = {}, fpMap = {};
    for (var ai = 0; ai < allRows.length; ai++) {
      var ar = allRows[ai];
      totalAnual += ar.totalPagar;
      totalPagado += ar.pagado;
      if (ar.totalPagar > 0) { countPos++; sumPos += ar.totalPagar; }

      var c = ar.categoria || 'Sin categoría';
      if (!catMap[c]) catMap[c] = { nombre: c, total: 0, count: 0 };
      catMap[c].total += ar.totalPagar;
      catMap[c].count++;

      var f = ar.formaPago || 'Sin especificar';
      if (!fpMap[f]) fpMap[f] = { nombre: f, total: 0, count: 0 };
      fpMap[f].total += ar.totalPagar;
      fpMap[f].count++;
    }

    function sortAndTop(map, limit) {
      var arr = [];
      for (var k in map) arr.push(map[k]);
      arr.sort(function(a, b) { return b.total - a.total; });
      var top = arr.slice(0, limit);
      for (var i = 0; i < top.length; i++) {
        top[i].pct = totalAnual > 0 ? top[i].total / totalAnual : 0;
      }
      return top;
    }

    var ticketPromedio = countPos > 0 ? sumPos / countPos : 0;
    var anio = new Date().getFullYear();
    // Try to extract year from spreadsheet name
    var ssName = ss.getName();
    var ym = ssName.match(/(\d{4})/);
    if (ym) anio = parseInt(ym[1], 10);

    return {
      view: 'ingresos',
      anio: anio,
      rows: allRows.slice(-500).reverse(),
      totalRows: allRows.length,
      resumenMensual: monthSummaries,
      totalAnual: totalAnual,
      totalPagado: totalPagado,
      numOperaciones: numOps,
      ticketPromedio: ticketPromedio,
      topCategorias: sortAndTop(catMap, 10),
      topFormasPago: sortAndTop(fpMap, 10),
      headers: ['ID','Fecha','Paciente','Categoría','Producto','P.V.P.','Desc.','Cantidad','Total','Pagado','Forma Pago','Sucursal']
    };
  } catch(ex) {
    return { view: 'ingresos', error: ex.message, rows: [], resumenMensual: [],
             totalAnual: 0, totalPagado: 0, numOperaciones: 0, ticketPromedio: 0,
             topCategorias: [], topFormasPago: [], headers: [] };
  }
}

/* ══════════════════════════════════════════════════════════════
   FORMATOS — Configuración de formato numérico por área
   ══════════════════════════════════════════════════════════════ */
var FORMATOS_TAB = 'Formatos';

function setupFormatos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(FORMATOS_TAB);
  if (sh) return {ok:true, msg:'Formatos ya existe'};
  sh = ss.insertSheet(FORMATOS_TAB);
  // Headers
  sh.getRange(1,1,1,6).setValues([['Area','Escala','Moneda','Decimales','SimboloMoneda','Activo']]);
  sh.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f4f6');
  // Defaults por área
  var defaults = [
    ['Finanzas',    'Miles',    'MXN', 2, '$', true],
    ['Ingresos',    'Miles',    'MXN', 2, '$', true],
    ['Conciliacion','Completo', 'MXN', 2, '$', true],
    ['Operaciones', 'Miles',    'MXN', 2, '$', true]
  ];
  sh.getRange(2,1,defaults.length,6).setValues(defaults);
  // Validación dropdown Escala
  var escalaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Completo','Miles','Millones'], true).build();
  sh.getRange(2,2,defaults.length,1).setDataValidation(escalaRule);
  // Validación dropdown Moneda
  var monedaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['MXN','USD','EUR'], true).build();
  sh.getRange(2,3,defaults.length,1).setDataValidation(monedaRule);
  // Validación decimales
  var decRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['0','1','2','3','4'], true).build();
  sh.getRange(2,4,defaults.length,1).setDataValidation(decRule);
  sh.setFrozenRows(1);
  return {ok:true, msg:'Formatos creada con defaults'};
}

function readFormatos() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(FORMATOS_TAB);
    if (!sh) {
      setupFormatos();
      sh = ss.getSheetByName(FORMATOS_TAB);
    }
    var raw = sh.getDataRange().getValues();
    if (raw.length < 2) return {ok:true, formatos:[]};
    var formatos = [];
    for (var i = 1; i < raw.length; i++) {
      var r = raw[i];
      if (!String(r[0]||'').trim()) continue;
      formatos.push({
        area:     String(r[0]||''),
        escala:   String(r[1]||'Completo'),
        moneda:   String(r[2]||'MXN'),
        decimales:Number(r[3])||2,
        simbolo:  String(r[4]||'$'),
        activo:   r[5]===true||String(r[5]).toUpperCase()==='TRUE'
      });
    }
    return {ok:true, formatos:formatos};
  } catch(ex) {
    return {ok:false, error:ex.message, formatos:[]};
  }
}

function saveFormatos(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(FORMATOS_TAB);
    if (!sh) { setupFormatos(); sh = ss.getSheetByName(FORMATOS_TAB); }
    var formatos = data.formatos || [];
    if (!formatos.length) return {ok:false, error:'Sin datos'};
    // Limpiar y reescribir
    var lr = sh.getLastRow();
    if (lr > 1) sh.getRange(2,1,lr-1,6).clearContent();
    var rows = formatos.map(function(f) {
      return [f.area||'', f.escala||'Completo', f.moneda||'MXN',
              Number(f.decimales)||2, f.simbolo||'$', f.activo!==false];
    });
    sh.getRange(2,1,rows.length,6).setValues(rows);
    return {ok:true, saved:rows.length};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

/* ══════════════════════════════════════════════════════════════
   BD_INGRESOS — Sistema de captura de operaciones multi-línea
   ══════════════════════════════════════════════════════════════ */
var BD_INGRESOS_TAB = 'BD_Ingresos';
var BD_INGRESOS_HEADERS = ['OP','Linea','Fecha','Paciente','Categoria','Producto',
  'PVP','Descuento','Cantidad','TotalPagar','Pagado','MontoFactMes',
  'FormaPago','Facturacion','Conciliacion','Contabilidad',
  'Observaciones','Factura','Poliza','USMX','CicloAltaBaja','Sucursal','ArchivoURL'];

function setupBDIngresos() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var existing = null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === BD_INGRESOS_TAB) { existing = sheets[i]; break; }
  }
  if (existing) return {ok:true, msg:'BD_Ingresos ya existe', rows:existing.getLastRow()-1};
  var sh = ss.insertSheet(BD_INGRESOS_TAB, 0);
  sh.getRange(1,1,1,BD_INGRESOS_HEADERS.length).setValues([BD_INGRESOS_HEADERS]);
  sh.getRange(1,1,1,BD_INGRESOS_HEADERS.length).setFontWeight('bold').setBackground('#f3f4f6');
  sh.setFrozenRows(1);
  return {ok:true, msg:'BD_Ingresos creada', headers:BD_INGRESOS_HEADERS};
}

function _getNextOP(sheet) {
  var lr = sheet.getLastRow();
  if (lr < 2) return 'OP-00001';
  var lastOP = String(sheet.getRange(lr, 1).getValue() || '');
  var m = lastOP.match(/OP-(\d+)/);
  var next = m ? (parseInt(m[1], 10) + 1) : 1;
  return 'OP-' + String(next).padStart(5, '0');
}

var INGRESOS_FOLDER_FACTURAS = '1t8--HM1xymgqGyBbIsI2jhMVCgQUBm9n';
var INGRESOS_FOLDER_PAGOS    = '1D9H3nNIrkgg2wqJtKXzhuSLDH6hIUoPk';

function saveIngreso(payload) {
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
    }
    if (!sheet) {
      setupBDIngresos();
      sheets = ss.getSheets();
      for (var j = 0; j < sheets.length; j++) {
        if (sheets[j].getName() === BD_INGRESOS_TAB) { sheet = sheets[j]; break; }
      }
    }
    if (!sheet) return {ok:false, error:'No se pudo crear BD_Ingresos'};

    var opId = _getNextOP(sheet);
    var lineas = payload.lineas || [];
    if (!lineas.length) return {ok:false, error:'No hay productos en la operación'};

    var fecha     = payload.fecha || '';
    var paciente  = payload.paciente || '';
    var formaPago = payload.formaPago || '';
    var sucursal  = payload.sucursal || '';
    var moneda    = payload.moneda || 'MX';
    var ciclo     = payload.ciclo || '';
    var obs       = payload.observaciones || '';
    var factura   = payload.factura || '';
    var poliza    = payload.poliza || '';
    var facturacionChk  = payload.facturacion === true || payload.facturacion === 'true';
    var conciliacionChk = payload.conciliacion === true || payload.conciliacion === 'true';
    var contabilidadChk = payload.contabilidad === true || payload.contabilidad === 'true';

    function num(v) { var n = parseFloat(String(v||'').replace(/[$,]/g,'')); return isNaN(n)?0:n; }

    var rows = [];
    var totalOP = 0;
    for (var li = 0; li < lineas.length; li++) {
      var l = lineas[li];
      var pvp  = num(l.pvp);
      var descPct = num(l.descuento) / 100; // Desc viene como % (ej. 10 = 10%)
      var cant = num(l.cantidad) || 1;
      var totalPagar = pvp * cant * (1 - descPct);
      var pagado = num(l.pagado) || totalPagar; // si no se especifica, pagado = total
      totalOP += totalPagar;

      // OP,Linea,Fecha,Paciente,Categoria,Producto,PVP,Descuento,Cantidad,TotalPagar,
      // Pagado,MontoFactMes,FormaPago,Facturacion,Conciliacion,Contabilidad,
      // Observaciones,Factura,Poliza,USMX,CicloAltaBaja,Sucursal,ArchivoURL
      rows.push([
        opId, li+1, fecha, paciente,
        l.categoria||'', l.producto||'',
        pvp, num(l.descuento), cant, totalPagar,
        pagado, li===0 ? num(payload.montoFactMes) : 0,
        formaPago,
        li===0 ? facturacionChk : false,
        li===0 ? conciliacionChk : false,
        li===0 ? contabilidadChk : false,
        li===0 ? obs : '',
        li===0 ? factura : '',
        li===0 ? poliza : '',
        moneda, ciclo, sucursal,
        '' // ArchivoURL — se llena al subir PDF
      ]);
    }

    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);

    return {ok:true, op:opId, lineas:rows.length, total:totalOP};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function uploadIngresoPDF(opId, tipo, fileName, base64Data, mimeType) {
  try {
    var folderId = (tipo === 'pago') ? INGRESOS_FOLDER_PAGOS : INGRESOS_FOLDER_FACTURAS;
    if (!folderId) return {ok:false, error:'Carpeta de Drive no configurada para '+tipo};
    var folder = DriveApp.getFolderById(folderId);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'application/pdf', fileName);
    var file = folder.createFile(blob);
    file.setName(opId + '_' + fileName);
    var url = file.getUrl();

    // Actualizar ArchivoURL en BD_Ingresos para todas las filas de este OP
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
    }
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var urlCol = BD_INGRESOS_HEADERS.indexOf('ArchivoURL') + 1; // 1-indexed
      for (var ri = 1; ri < data.length; ri++) {
        if (String(data[ri][0]) === opId) {
          sheet.getRange(ri+1, urlCol).setValue(url);
        }
      }
    }
    return {ok:true, url:url, fileName:opId+'_'+fileName};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function migrateIngresosToDBD() {
  // Destino: BD_Ingresos en el spreadsheet 2026
  var ssDest = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sheet = null;
  var destSheets = ssDest.getSheets();
  for (var i = 0; i < destSheets.length; i++) {
    if (destSheets[i].getName() === BD_INGRESOS_TAB) { sheet = destSheets[i]; break; }
  }
  if (!sheet) {
    setupBDIngresos();
    destSheets = ssDest.getSheets();
    for (var j = 0; j < destSheets.length; j++) {
      if (destSheets[j].getName() === BD_INGRESOS_TAB) { sheet = destSheets[j]; break; }
    }
  }

  var MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var opCounter = 0;
  var lr = sheet.getLastRow();
  if (lr > 1) {
    var lastOP = String(sheet.getRange(lr,1).getValue()||'');
    var m = lastOP.match(/OP-(\d+)/);
    if (m) opCounter = parseInt(m[1],10);
  }

  function num(v){var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}
  function dt(v){
    if(!v) return '';
    if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
    return String(v);
  }

  // Procesar múltiples spreadsheets en orden cronológico
  var sources = [
    {id: INGRESOS_SS_2024, label: '2024'},
    {id: INGRESOS_SS_2025, label: '2025'},
    {id: INGRESOS_SS_ID,   label: '2026'}
  ];

  var totalMigrated = 0;
  var log = [];

  for (var si = 0; si < sources.length; si++) {
    var srcId = sources[si].id;
    if (!srcId) continue;
    var srcSS;
    try { srcSS = SpreadsheetApp.openById(srcId); }
    catch(ex) { log.push(sources[si].label + ': Error abriendo — ' + ex.message); continue; }

    var srcSheets = srcSS.getSheets();
    var yearMigrated = 0;

    // Procesar meses en orden (Enero primero)
    for (var mi = 0; mi < MESES.length; mi++) {
      var tabName = MESES[mi];
      var srcSheet = null;
      for (var ti = 0; ti < srcSheets.length; ti++) {
        if (srcSheets[ti].getName().trim() === tabName) { srcSheet = srcSheets[ti]; break; }
      }
      if (!srcSheet) continue;

      var raw = srcSheet.getDataRange().getValues();
      if (raw.length < 2) continue;

      // Agrupar filas por ID original (columna A)
      var groups = {};
      var groupOrder = [];
      for (var ri = 1; ri < raw.length; ri++) {
        var r = raw[ri];
        var origId = String(r[0]||'').trim();
        if (!origId || origId === '0') continue;
        var idNum = Number(origId);
        if (isNaN(idNum) || idNum <= 0) continue;
        if (!groups[origId]) { groups[origId] = []; groupOrder.push(origId); }
        groups[origId].push(r);
      }

      var batchRows = [];
      for (var gi = 0; gi < groupOrder.length; gi++) {
        var grp = groups[groupOrder[gi]];
        opCounter++;
        var opId = 'OP-' + String(opCounter).padStart(5,'0');
        for (var li = 0; li < grp.length; li++) {
          var r = grp[li];
          var pvp  = num(r[5]);
          var descPct = num(r[6]);
          var cant = num(r[7]) || 1;
          var totalPagar = num(r[8]) || (pvp * cant * (1 - descPct/100));
          var pagado = num(r[9]);
          var montoFact = num(r[10]);
          var facChk  = r[12]===true||String(r[12]).toUpperCase()==='TRUE';
          var conChk  = r[13]===true||String(r[13]).toUpperCase()==='TRUE';
          var ctaChk  = r[14]===true||String(r[14]).toUpperCase()==='TRUE';
          batchRows.push([
            opId, li+1, dt(r[1]), String(r[2]||''),
            String(r[3]||''), String(r[4]||''),
            pvp, descPct, cant, totalPagar,
            pagado, montoFact,
            String(r[11]||''),
            facChk, conChk, ctaChk,
            String(r[15]||''), String(r[16]||''), String(r[17]||''),
            String(r[18]||''), String(r[19]||''),
            (r.length > 20 ? String(r[20]||'') : ''),
            ''
          ]);
        }
      }

      if (batchRows.length) {
        sheet.getRange(sheet.getLastRow()+1, 1, batchRows.length, batchRows[0].length).setValues(batchRows);
        totalMigrated += batchRows.length;
        yearMigrated += batchRows.length;
      }
    }
    log.push(sources[si].label + ': ' + yearMigrated + ' filas migradas');
  }

  return {ok:true, totalMigrated:totalMigrated, lastOP:'OP-'+String(opCounter).padStart(5,'0'), log:log};
}

/* ══════════════════════════════════════════════════════════════
   fixOPGrouping — Reagrupa OP por paciente+fecha
   Ejecutar UNA VEZ desde el editor de Apps Script (Run ▶)
   ══════════════════════════════════════════════════════════════ */
function fixOPGrouping() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sheet = null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
  }
  if (!sheet) return {ok:false, error:'BD_Ingresos no encontrada'};

  var raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return {ok:false, error:'Sin datos'};

  function dt(v) {
    if (!v) return '';
    if (v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
    return String(v).substring(0,10);
  }

  // Paso 1: Leer todas las filas de datos (skip header), eliminar vacías
  var dataRows = [];
  for (var ri = 1; ri < raw.length; ri++) {
    var r = raw[ri];
    var paciente = String(r[3]||'').trim();
    var fecha = dt(r[2]);
    // Saltar filas vacías (sin paciente Y sin fecha Y sin producto)
    if (!paciente && !fecha && !String(r[5]||'').trim()) continue;
    dataRows.push(r);
  }

  // Paso 2: Ordenar por fecha ASC, luego paciente
  dataRows.sort(function(a, b) {
    var dA = dt(a[2]), dB = dt(b[2]);
    if (dA < dB) return -1;
    if (dA > dB) return 1;
    var pA = String(a[3]||'').toLowerCase(), pB = String(b[3]||'').toLowerCase();
    if (pA < pB) return -1;
    if (pA > pB) return 1;
    return 0;
  });

  // Paso 3: Agrupar por paciente+fecha → mismo OP
  var opCounter = 0;
  var lastKey = '';
  var lastOP = '';
  var lineInOP = 0;
  var newRows = [];

  for (var di = 0; di < dataRows.length; di++) {
    var r = dataRows[di];
    var paciente = String(r[3]||'').trim();
    var fecha = dt(r[2]);
    var key = (paciente + '||' + fecha).toLowerCase();

    if (key !== lastKey || !lastKey) {
      // Nueva operación
      opCounter++;
      lastOP = 'OP-' + String(opCounter).padStart(5, '0');
      lastKey = key;
      lineInOP = 1;
    } else {
      lineInOP++;
    }

    // Construir fila nueva con OP y Linea corregidos
    var newRow = [lastOP, lineInOP];
    for (var ci = 2; ci < r.length; ci++) {
      newRow.push(r[ci]);
    }
    // Asegurar 23 columnas
    while (newRow.length < 23) newRow.push('');
    newRows.push(newRow);
  }

  // Paso 4: Reescribir BD_Ingresos (preservar header)
  var header = raw[0];
  var lr = sheet.getLastRow();
  if (lr > 1) sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).clearContent();

  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  // Contar operaciones únicas
  var uniqueOps = {};
  newRows.forEach(function(r) { uniqueOps[r[0]] = true; });

  return {
    ok: true,
    filasOriginales: raw.length - 1,
    filasVaciasEliminadas: (raw.length - 1) - dataRows.length,
    filasFinales: newRows.length,
    operacionesUnicas: Object.keys(uniqueOps).length,
    ultimoOP: lastOP
  };
}

/* ── Constantes definidas en api_config.gs (mismo proyecto GAS) ──
   SHEET_ID, AUTH_SECRET, ER_SS_ID, BANKS_SS_ID, LAB_SS_ID, MED_SS_ID,
   PAC_SS_ID, PROD_SS_ID, QX_SS_ID, ER_GID, BUDGET_GID, PL_GID,
   BANKS_GID, CAPTURA_SHEETS, SHEET_ALIASES, CAPTURA_SHEET_ID_DEFAULT
   ────────────────────────────────────────────────────────────────── */

/* ── Autenticación: helpers ──────────────────────────────────── */