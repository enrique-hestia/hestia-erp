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
    if (body.action === 'login') {
      return jsonResponse(handleLogin(body.email || '', body.password || ''));
    }
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
    if (body.action === 'uploadFile') {
      return jsonResponse(uploadFile(body));
    }
    if (body.action === 'uploadIngresoPDF') {
      return jsonResponse(uploadIngresoPDF(body.opId, body.tipo||'factura', body.fileName, body.base64, body.mimeType));
    }
    if (body.action === 'updateIngresoFiscal') {
      return jsonResponse(updateIngresoFiscal(body));
    }
    if (body.action === 'setupBDIngresos') {
      return jsonResponse(setupBDIngresos());
    }
    if (body.action === 'setupBDProductos') {
      return jsonResponse(setupBDProductos());
    }
    if (body.action === 'saveNewProducto') {
      return jsonResponse(saveNewProducto(body));
    }
    if (body.action === 'saveProductoPrecio') {
      return jsonResponse(saveProductoPrecio(body.productoId, body.precio, body.vigencia, body.usuario));
    }
    if (body.action === 'saveCxP') {
      return jsonResponse(saveCxP(body));
    }
    if (body.action === 'deleteCxPRow')            return jsonResponse(deleteCxPRow(body));
    if (body.action === 'updateCxP') {
      return jsonResponse(updateCxP(body));
    }
    if (body.action === 'pagarCxP') {
      return jsonResponse(pagarCxP(body));
    }
    // Gastos fijos (recurrentes)
    if (body.action === 'gfAll')                   return jsonResponse({ok:true, propuestas:readGastosFijosPropuestas(body.periodo||''), catalogo:readGastosFijos()});
    if (body.action === 'setupGastosFijos')        return jsonResponse(setupGastosFijos());
    if (body.action === 'saveGastoFijo')           return jsonResponse(saveGastoFijo(body));
    if (body.action === 'saveGastosFijosBatch')    return jsonResponse(saveGastosFijosBatch(body));
    if (body.action === 'reconstruirCatalogoGF')   return jsonResponse(reconstruirCatalogoGF(body));
    if (body.action === 'updateGastoFijo')         return jsonResponse(updateGastoFijo(body));
    if (body.action === 'toggleGastoFijo')         return jsonResponse(toggleGastoFijo(body));
    if (body.action === 'deleteGastoFijo')         return jsonResponse(deleteGastoFijo(body));
    if (body.action === 'programarGastoFijo')      return jsonResponse(programarGastoFijo(body));
    if (body.action === 'programarGastosFijosBatch') return jsonResponse(programarGastosFijosBatch(body));
    if (body.action === 'regresarCxPaProgramacion') {
      if (typeof regresarCxPaProgramacion !== 'function')
        return jsonResponse({ok:false, error:'Actualiza gastosfijos.gs en Apps Script y redespliega.'});
      return jsonResponse(regresarCxPaProgramacion(body));
    }
    if (body.action === 'bulkUpdateCxPMonto')         return jsonResponse(bulkUpdateCxPMonto(body));
    if (body.action === 'conciliaAMEX')              return jsonResponse(conciliaAMEX(body));
    if (body.action === 'amexCorte')                 return jsonResponse(readAmexCorte(body));
    if (body.action === 'estadoCuenta') {
      if (typeof readEstadoCuentaPaciente !== 'function')
        return jsonResponse({ok:false, error:'Agrega analisis.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readEstadoCuentaPaciente(body.paciente||''));
    }
    // Tareas programadas (scheduler)
    if (body.action === 'updateScheduledTask')     return jsonResponse(updateScheduledTask(body));
    if (body.action === 'setupScheduledTriggers')  return jsonResponse(setupScheduledTriggers());
    // Recordatorios (agenda personal) — abierto a cualquier área
    if (body.action === 'setupRecordatorios')      return jsonResponse(setupRecordatorios());
    if (body.action === 'saveRecordatorio')        return jsonResponse(saveRecordatorio(body));
    if (body.action === 'updateRecordatorioEstado') return jsonResponse(updateRecordatorioEstado(body));
    if (body.action === 'updateRecordatorio') {
      if (typeof updateRecordatorio !== 'function')
        return jsonResponse({ok:false, error:'Actualiza recordatorios.gs en Apps Script y redespliega.'});
      return jsonResponse(updateRecordatorio(body));
    }
    if (body.action === 'deleteRecordatorio')      return jsonResponse(deleteRecordatorio(body));
    if (body.action === 'updateProducto') {
      return jsonResponse(updateProducto(body));
    }
    if (body.action === 'createProducto') {
      return jsonResponse(createProducto(body));
    }
    if (body.action === 'exportarCatalogoProductos') {
      return jsonResponse(exportarCatalogoProductos(body.usuario));
    }
    if (body.action === 'vincularComprobanteEgreso') {
      if (typeof vincularComprobanteEgreso !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(vincularComprobanteEgreso(body));
    }
    if (body.action === 'configurarMenuComprobantes') {
      if (typeof configurarMenuComprobantes !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(configurarMenuComprobantes());
    }
    if (body.action === 'buscarXmlParaEgreso') {
      if (typeof buscarXmlParaEgreso !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(buscarXmlParaEgreso(body));
    }
    if (body.action === 'vincularComprobantesLote') {
      if (typeof vincularComprobantesLote !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(vincularComprobantesLote(body));
    }
    if (body.action === 'enviarDocumentoCorreo') {
      if (typeof enviarDocumentoCorreo !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(enviarDocumentoCorreo(body));
    }
    if (body.action === 'saveSummaryConfig') {
      if (typeof saveSummaryConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega summary.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(saveSummaryConfig(body));
    }
    if (body.action === 'setupSummaryConfig') {
      if (typeof setupSummaryConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega summary.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(setupSummaryConfig());
    }
    if (body.action === 'saveBoardConfig') {
      if (typeof saveBoardConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega board.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(saveBoardConfig(body));
    }
    if (body.action === 'guardarClasifProveedor') {
      if (typeof guardarClasifProveedor !== 'function')
        return jsonResponse({ok:false, error:'Agrega prov_defaults.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(guardarClasifProveedor(body));
    }
    if (body.action === 'savePlantillaCorreo') {
      if (typeof savePlantillaCorreo !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(savePlantillaCorreo(body));
    }
    if (body.action === 'importarCatalogoProductosBatch') {
      return jsonResponse(importarCatalogoProductosBatch(body));
    }
    if (body.action === 'vincularXmlFactura') {
      return jsonResponse(vincularXmlFactura(body));
    }
    if (body.action === 'uploadYVincularXmlFactura') {
      return jsonResponse(uploadYVincularXmlFactura(body));
    }
    if (body.action === 'vincularAutomaticoLote') {
      return jsonResponse(vincularAutomaticoLote(body.fechaInicio, body.fechaFin, body.usuario));
    }
    if (body.action === 'setupInventarioMedicamentos') {
      return jsonResponse(setupInventarioMedicamentos());
    }
    if (body.action === 'configurarMenuMedicamentos') {
      return jsonResponse(configurarMenuMedicamentos());
    }
    if (body.action === 'setupCxPCreditosAbonos') {
      if (typeof setupCxPCreditosAbonos !== 'function')
        return jsonResponse({ok:false, error:'Agrega cxp_creditos.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(setupCxPCreditosAbonos());
    }
    if (body.action === 'aplicarAbonoCxP') {
      if (typeof aplicarAbonoCxP !== 'function')
        return jsonResponse({ok:false, error:'Agrega cxp_creditos.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(aplicarAbonoCxP(body));
    }
    if (body.action === 'cancelarOrdenCxP') {
      if (typeof cancelarOrdenCxP !== 'function')
        return jsonResponse({ok:false, error:'Agrega cxp_creditos.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(cancelarOrdenCxP(body));
    }
    if (body.action === 'revertirAbonosDeOrden') {
      if (typeof revertirAbonosDeOrden !== 'function')
        return jsonResponse({ok:false, error:'Actualiza cxp_creditos.gs en Apps Script y redespliega.'});
      return jsonResponse(revertirAbonosDeOrden(body));
    }
    if (body.action === 'repararAbonosCruzados') {
      if (typeof repararAbonosCruzados !== 'function')
        return jsonResponse({ok:false, error:'Actualiza cxp_creditos.gs en Apps Script y redespliega.'});
      return jsonResponse(repararAbonosCruzados());
    }
    if (body.action === 'saveMedicamento') {
      return jsonResponse(saveMedicamento(body));
    }
    if (body.action === 'updateMedicamento') {
      return jsonResponse(updateMedicamento(body));
    }
    if (body.action === 'ajustarInventarioMedicamento') {
      return jsonResponse(ajustarInventarioMedicamento(body));
    }
    if (body.action === 'saveCombo') {
      return jsonResponse(saveCombo(body));
    }
    if (body.action === 'eliminarCombo') {
      return jsonResponse(eliminarCombo(body));
    }
    if (body.action === 'setupOrdenesCompra') {
      return jsonResponse(setupOrdenesCompra());
    }
    if (body.action === 'crearOrdenCompra') {
      return jsonResponse(crearOrdenCompra(body));
    }
    if (body.action === 'marcarOrdenRecibida') {
      return jsonResponse(marcarOrdenRecibida(body));
    }
    if (body.action === 'registrarSobranteInventario') {
      return jsonResponse(registrarSobranteInventario(body));
    }
    if (body.action === 'migrarHistorialInventario') {
      return jsonResponse(migrarHistorialInventario(body.confirmar === true));
    }
    if (body.action === 'generarReporteContaDigital') {
      return jsonResponse(generarReporteContaDigital(body.fechaInicio, body.fechaFin, body.usuario));
    }
    if (body.action === 'generarReporteContaDigitalPendientes') {
      return jsonResponse(generarReporteContaDigitalPendientes(body.fechaInicio, body.fechaFin, body.usuario));
    }
    if (body.action === 'aplicarDatosFiscalesPacientes') {
      return jsonResponse(aplicarDatosFiscalesPacientes(body));
    }
    if (body.action === 'backfillRazonSocialDesdeXml') {
      return jsonResponse(backfillRazonSocialDesdeXml(body.fechaInicio, body.fechaFin, body.usuario));
    }
    if (body.action === 'saveLista') {
      return jsonResponse(saveLista(body));
    }
    if (body.action === 'setupProveedores') {
      return jsonResponse(setupProveedores());
    }
    if (body.action === 'saveProveedor') {
      return jsonResponse(saveProveedor(body));
    }
    if (body.action === 'updateProveedor') {
      return jsonResponse(updateProveedor(body));
    }
    if (body.action === 'autocompletarProveedores') {
      if (typeof autocompletarProveedoresDesdeXML !== 'function')
        return jsonResponse({ok:false, error:'Actualiza providers.gs en Apps Script y redespliega.'});
      return jsonResponse(autocompletarProveedoresDesdeXML(body));
    }
    if (body.action === 'aplicarAutocompletarProveedores') {
      if (typeof aplicarAutocompletarProveedores !== 'function')
        return jsonResponse({ok:false, error:'Actualiza providers.gs en Apps Script y redespliega.'});
      return jsonResponse(aplicarAutocompletarProveedores(body));
    }
    if (body.action === 'setupPresupuesto') {
      return jsonResponse(setupPresupuesto());
    }
    if (body.action === 'savePresupuestoMeta') {
      return jsonResponse(savePresupuestoMeta(body));
    }
    if (body.action === 'savePresupuestoMetasBatch') {
      return jsonResponse(savePresupuestoMetasBatch(body));
    }
    if (body.action === 'saveGruposPresupuesto') {
      if (typeof saveGruposPresupuesto !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(saveGruposPresupuesto(body));
    }
    if (body.action === 'saveMenu') {
      if (typeof saveMenu !== 'function')
        return jsonResponse({ok:false, error:'Actualiza core.gs en Apps Script y redespliega.'});
      return jsonResponse(saveMenu(body));
    }
    if (body.action === 'presSetLock') {
      if (typeof presSetLock !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(presSetLock(body));
    }
    if (body.action === 'presSetEscenarios') {
      if (typeof presSetEscenarios !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(presSetEscenarios(body));
    }
    if (body.action === 'factVencidaGuardar') {
      if (typeof factVencidaGuardar !== 'function')
        return jsonResponse({ok:false, error:'Agrega devengado.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(factVencidaGuardar(body));
    }
    if (body.action === 'chatSend') {
      if (typeof chatSend !== 'function')
        return jsonResponse({ok:false, error:'Agrega chat.gs en Apps Script y redespliega.'});
      return jsonResponse(chatSend(body));
    }
    if (body.action === 'chatPoll') {
      if (typeof chatPoll !== 'function')
        return jsonResponse({ok:false, error:'Agrega chat.gs en Apps Script y redespliega.'});
      return jsonResponse(chatPoll(body));
    }
    if (body.action === 'chatUpload') {
      if (typeof chatUpload !== 'function')
        return jsonResponse({ok:false, error:'Agrega chat.gs en Apps Script y redespliega.'});
      return jsonResponse(chatUpload(body));
    }
    if (body.action === 'presSetPeriodoAbierto') {
      if (typeof presSetPeriodoAbierto !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(presSetPeriodoAbierto(body));
    }
    if (body.action === 'saveSemanalConfig') {
      if (typeof saveSemanalConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega semanal.gs en Apps Script y redespliega.'});
      return jsonResponse(saveSemanalConfig(body));
    }
    if (body.action === 'updateProductoSKU') {
      return jsonResponse(updateProductoSKU(body.productoId, body.sku, body.usuario));
    }
    if (body.action === 'updateProductoID') {
      return jsonResponse(updateProductoID(body.productoIdViejo, body.productoIdNuevo, body.usuario));
    }
    if (body.action === 'saveEgreso') {
      return jsonResponse(saveEgreso(body));
    }
    if (body.action === 'uploadEgresoPDF') {
      return jsonResponse(uploadEgresoPDF(body));
    }
    if (body.action === 'updateEgresoField') {
      return jsonResponse(updateEgresoField(body));
    }
    if (body.action === 'repararEgresosSinFecha') {
      return jsonResponse(repararEgresosSinFecha(body));
    }
    if (body.action === 'guardarReferenciaEgreso') {
      return jsonResponse(guardarReferenciaEgreso(body));
    }
    if (body.action === 'deleteEgreso') {
      return jsonResponse(deleteEgreso(body));
    }
    if (body.action === 'saveDropdown') {
      return jsonResponse(saveDropdownValues(body));
    }
    if (body.action === 'updateIngreso') {
      return jsonResponse(updateIngreso(body));
    }
    if (body.action === 'updateIngresoConBancos') {
      return jsonResponse(updateIngresoConBancos(body));
    }
    if (body.action === 'deleteIngreso') {
      return jsonResponse(deleteIngreso(body));
    }
    if (body.action === 'renamePacienteIngresos') {
      return jsonResponse(renamePacienteIngresos(body.oldNombre, body.newNombre));
    }
    if (body.action === 'updateCajaChica') {
      return jsonResponse(updateCajaChicaRow(body));
    }
    if (body.action === 'saveCajaChicaIngreso') {
      return jsonResponse(saveCajaChicaIngreso(body));
    }
    if (body.action === 'saveuser') {
      var ss      = SpreadsheetApp.openById(SHEET_ID);
      var tkEmail = verifyToken(body.token || '');
      if (!tkEmail) return jsonResponse({ error: 'Sesión inválida.', code: 401 });
      var cu = getUserRow(ss, tkEmail);
      if (!cu || String(cu.rol||'').toLowerCase() !== 'admin') return jsonResponse({ error: 'Sin permisos de administrador.' });
      var shU  = ss.getSheetByName('Usuarios');
      if (!shU) return jsonResponse({ error: 'Hoja Usuarios no encontrada.' });
      var hdrs = shU.getRange(1,1,1,shU.getLastColumn()).getValues()[0];
      var rowNum = parseInt(body.rowNum || '0');
      var newRow = hdrs.map(function(h, i) {
        var key = String(h).trim();
        if (key === 'Contraseña' && !body[key] && rowNum > 1) {
          return shU.getRange(rowNum, i+1).getValue();
        }
        return body[key] !== undefined ? body[key] : '';
      });
      if (rowNum > 1) {
        shU.getRange(rowNum, 1, 1, hdrs.length).setValues([newRow]);
      } else {
        shU.appendRow(newRow);
      }
      return jsonResponse({ success: true });
    }
    if (body.action === 'reporteProveedor')    return jsonResponse(readReporteProveedor(body));
    if (body.action === 'setupEgresosAnio')    return jsonResponse(setupEgresosAnio(body.anio));
    return jsonResponse({error:'Accion desconocida: ' + body.action});
  } catch(ex) { return jsonResponse({error: ex.message}); }
}

function readBanksData(fechaInicio, fechaFin) {
  try {
    var ss = SpreadsheetApp.openById(BANKS_SS_ID);
    var sh = ss.getSheets();
    function byGid(gid){ for(var i=0;i<sh.length;i++) if(sh[i].getSheetId()===gid) return sh[i]; return null; }
    function num(v){ if(typeof v==='number') return v; var n=parseFloat(String(v||'').replace(/[$,\s]/g,'')); return isNaN(n)?0:n; }
    function dt(v){ if(!v) return ''; if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0'); return String(v); }

    // Filtro por rango de fechas (YYYY-MM-DD). Si no viene, muestra todo.
    var fi = fechaInicio || '';
    var ff = fechaFin    || '';
    function inRange(f){ if(!fi&&!ff) return true; if(!f) return true; return (!fi||f>=fi) && (!ff||f<=ff); }

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
      }).filter(function(m){ return inRange(m.fecha); }).reverse();
      return B;
    }
    function rAmex(sheet) {
      var B={id:'amex',nombre:'AMEX',color:'#007bc1',saldo:0,movimientos:[],totalRows:0};
      if(!sheet) return B;
      var r=sheet.getDataRange().getValues(); if(r.length<2) return B;
      for(var i=r.length-1;i>=1;i--){ var s=num(r[i][2]); if(s!==0){B.saldo=s;break;} }
      B.totalRows=r.length-1;
      var amexRows=[];
      for(var i=1;i<r.length;i++){
        var m=num(r[i][1]);
        var f=dt(r[i][0]);
        if(!f&&m===0) continue; // saltar filas completamente vacías
        if(!inRange(f)) continue; // respetar filtro de fechas
        amexRows.push({rowNum:i+1,fecha:f,monto:m,saldo:num(r[i][2]),referencia:String(r[i][3]||''),
                       usd:num(r[i][4]),tipoCambio:num(r[i][5]),notas:String(r[i][6]||''),
                       poliza:String(r[i][7]||''),mes:String(r[i][8]||''),tipo:m>=0?'cargo':'pago'});
      }
      B.movimientos=amexRows.reverse(); // más recientes primero
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
      B.movimientos=dataRows.filter(function(entry){
        return inRange(dt(entry.row[1]));
      }).reverse().map(function(entry){
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
        }).filter(function(m){ return inRange(m.fecha); }).reverse();
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
   BD_CxP — Cuentas por Pagar separadas de Egresos
   ══════════════════════════════════════════════════════════════ */
var BD_CXP_TAB = 'BD_CxP';
var BD_CXP_HEADERS = ['ID','Mes','Prioridad','Proveedor','Contable','Tipo','Subtipo','Concepto',
  'Monto','Notas','Vencimiento','Facturacion','Pagado','Contabilidad','Poliza',
  'FormaPago','Observaciones','LinkFactura','LinkPago','FechaRegistro'];

function setupBDCxP() {
  var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
  var sh = ss.getSheetByName(BD_CXP_TAB);
  if (!sh) {
    sh = ss.insertSheet(BD_CXP_TAB);
    sh.getRange(1,1,1,BD_CXP_HEADERS.length).setValues([BD_CXP_HEADERS]);
    sh.getRange(1,1,1,BD_CXP_HEADERS.length).setFontWeight('bold').setBackground('#fef3c7');
    sh.setFrozenRows(1);
  }
  return {ok:true, msg:'BD_CxP creada'};
}

function _getNextCxPID(sheet) {
  var lr = sheet.getLastRow();
  if (lr < 2) return 'CXP-00001';
  var last = String(sheet.getRange(lr,1).getValue()||'');
  var m = last.match(/CXP-(\d+)/);
  return 'CXP-' + String((m ? parseInt(m[1],10) : 0) + 1).padStart(5,'0');
}

// Escribe Divisa + (si USD) Monto USD y Tipo de Cambio en la fila de la CxP.
function _cxpGuardarDivisa(sh, rowNum, body) {
  var div = String(body.divisa||'MXN').toUpperCase()==='USD' ? 'USD' : 'MXN';
  var iDiv = _egColEnsure(sh, 'divisa', 'Divisa');
  sh.getRange(rowNum, iDiv).setValue(div);
  var iUSD = _egColEnsure(sh, 'usd', 'Monto USD');
  var iTC  = _egColEnsure(sh, 'tipo de cambio', 'Tipo de Cambio');
  if (div === 'USD') {
    sh.getRange(rowNum, iUSD).setValue(parseFloat(body.montoUSD)||0);
    sh.getRange(rowNum, iTC).setValue(parseFloat(body.tipoCambio)||0);
  } else {
    sh.getRange(rowNum, iUSD).setValue('');
    sh.getRange(rowNum, iTC).setValue('');
  }
}

function saveCxP(body) {
  // Guarda directamente en Egresos2026 SIN fecha (col B vacía = CxP)
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var sh = ss.getSheetByName(egTab);
    if (!sh) return {ok:false, error:'Hoja Egresos no encontrada'};
    // Siguiente ID consecutivo
    var lr = sh.getLastRow();
    var lastId = 0;
    if (lr > 1) {
      var ids = sh.getRange(2, 1, lr-1, 1).getValues();
      for (var i=0;i<ids.length;i++) { var n=parseInt(ids[i][0]); if(n>lastId) lastId=n; }
    }
    var newId = lastId + 1;
    var monto = parseFloat(String(body.monto||'').replace(/[$,]/g,''))||0;
    // Egresos row: ID, Fecha(vacía), Mes, col1, Proveedor, Contable, Tipo, Subtipo, Concepto, Monto,
    //              Notas, Vencimiento, Facturacion, Pagado, Contabilidad, Poliza, FormaPago, Obs, LinkFact, LinkPago
    var row = [
      newId, '', body.mes||'', body.prioridad||1,
      body.proveedor||'', body.contable||'Gasto', body.tipo||'Variable', body.subtipo||'',
      body.concepto||'', monto, body.notas||'',
      body.vencimiento||'',
      body.facturacion===true, false, false,
      body.poliza||'', '', body.observaciones||'',
      body.linkFactura||'', ''
    ];
    sh.appendRow(row);
    var newRowNum = sh.getLastRow();
    // Cotización adjunta al registrar (link ya subido a Drive por uploadFile)
    if (body.linkCotizacion) {
      try {
        var iCot = _egColEnsure(sh, 'cotiz', 'Link Cotizacion');
        sh.getRange(newRowNum, iCot).setRichTextValue(
          SpreadsheetApp.newRichTextValue().setText('Cotización').setLinkUrl(body.linkCotizacion).build());
      } catch(_c) { /* no romper el guardado si falla el link */ }
    }
    // Divisa + tipo de cambio (CxP en moneda extranjera). Monto (col J) se guarda ya en MXN.
    try { _cxpGuardarDivisa(sh, newRowNum, body); } catch(_d) {}
    logAudit(body.usuario||'sistema','CxP','Crear',String(newId),'','',body.proveedor+' | $'+monto+' | '+(String(body.divisa||'MXN').toUpperCase())+' | Vence: '+body.vencimiento);
    return {ok:true, id:newId, monto:monto, rowNum:newRowNum};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function readBDCxP() {
  // Lee directamente de Egresos2026: filas sin fecha (col B) y no pagadas = CxP
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var sh = ss.getSheetByName(egTab);
    if (!sh) return {ok:true, rows:[], resumen:{vencido:0,hoy:0,semana:0,mes:0,totalVencido:0,totalHoy:0,totalSemana:0,totalMes:0,totalPendiente:0}};
    var raw = sh.getDataRange().getValues();
    if (raw.length < 2) return {ok:true, rows:[], resumen:{vencido:0,hoy:0,semana:0,mes:0,totalVencido:0,totalHoy:0,totalSemana:0,totalMes:0,totalPendiente:0}};

    var hoy = new Date(); hoy.setHours(0,0,0,0);
    function dt(v){if(!v)return'';if(v instanceof Date)return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');return String(v);}
    function parseD(v){if(!v)return null;if(v instanceof Date){var d=new Date(v);d.setHours(0,0,0,0);return d;}var s=String(v).trim();var d=new Date(s);return isNaN(d)?null:d;}
    function num(v){var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}

    // Columnas Egresos2026: A=ID(0), B=Fecha(1), C=Mes(2), D=col1(3), E=Proveedor(4),
    // F=Contable(5), G=Tipo(6), H=Subtipo(7), I=Concepto(8), J=Egresos(9),
    // K=Notas(10), L=Vencimiento(11), M=Facturacion(12), N=Pagado(13),
    // O=Contabilidad(14), P=Poliza(15), Q=FormaPago(16), R=Obs(17), S=LinkFact(18), T=LinkPago(19)

    var resumen = {vencido:0,hoy:0,semana:0,mes:0,totalVencido:0,totalHoy:0,totalSemana:0,totalMes:0,totalPendiente:0};
    var rows = [];
    var iCotiz = -1, iDiv = -1, iEstatus = -1, iRec = -1, iUSD = -1, iTC = -1, hdr0 = raw[0]||[];
    for (var hc=0; hc<hdr0.length; hc++){
      var h0=String(hdr0[hc]).toLowerCase();
      if(iCotiz<0 && h0.indexOf('cotiz')>-1) iCotiz=hc;
      if(iDiv<0 && h0.indexOf('divisa')>-1) iDiv=hc;
      if(iEstatus<0 && h0.indexOf('estatus')>-1) iEstatus=hc;
      if(iRec<0 && h0.indexOf('recurrente')>-1) iRec=hc;
      if(iUSD<0 && h0.indexOf('usd')>-1) iUSD=hc;
      if(iTC<0 && h0.indexOf('tipo de cambio')>-1) iTC=hc;
    }

    for (var i=1;i<raw.length;i++) {
      var r = raw[i];
      var egId = String(r[0]||'').trim();
      if (!egId) continue;
      var fecha = r[1]; // col B = fecha de pago
      var pagado = r[13]===true||String(r[13]).toUpperCase()==='TRUE';
      var estatusRow = iEstatus>-1 ? String(r[iEstatus]||'').trim() : '';
      if (estatusRow === 'Cancelada') continue; // cancelada: no es CxP pendiente ni egreso
      // CxP = sin fecha de pago Y no pagado
      if (fecha || pagado) continue;
      var venc = parseD(r[11]); // col L = vencimiento
      var monto = num(r[9]);
      var dias = venc ? Math.round((venc-hoy)/86400000) : null;
      var urgencia = dias===null?'sin-fecha':dias<0?'vencido':dias===0?'hoy':dias<=7?'semana':dias<=30?'mes':'ok';

      if(urgencia==='vencido'){resumen.vencido++;resumen.totalVencido+=monto;}
      else if(urgencia==='hoy'){resumen.hoy++;resumen.totalHoy+=monto;}
      else if(urgencia==='semana'){resumen.semana++;resumen.totalSemana+=monto;}
      else if(urgencia==='mes'){resumen.mes++;resumen.totalMes+=monto;}
      resumen.totalPendiente+=monto;

      rows.push({
        rowNum:i+1, id:egId, mes:String(r[2]||''),
        proveedor:String(r[4]||''), contable:String(r[5]||''), tipo:String(r[6]||''),
        subtipo:String(r[7]||''), concepto:String(r[8]||''), monto:monto,
        notas:String(r[10]||''), vencimiento:dt(r[11]),
        facturacion:r[12]===true||String(r[12]).toUpperCase()==='TRUE',
        pagado:false, contabilidad:r[14]===true||String(r[14]).toUpperCase()==='TRUE',
        poliza:String(r[15]||''), formaPago:String(r[16]||''),
        observaciones:String(r[17]||''), linkFactura:String(r[18]||''), linkFacturaUrl:'',
        linkPago:String(r[19]||''), linkPagoUrl:'',
        linkCotizacion: iCotiz>-1 ? String(r[iCotiz]||'') : '', linkCotizacionUrl:'',
        divisa: (iDiv>-1 && String(r[iDiv]||'').toUpperCase()==='USD') ? 'USD' : 'MXN',
        montoUSD: iUSD>-1 ? num(r[iUSD]) : 0,
        tipoCambio: iTC>-1 ? num(r[iTC]) : 0,
        recurrenteId: iRec>-1 ? String(r[iRec]||'').trim() : '',
        dias:dias, urgencia:urgencia
      });
    }
    // Hipervínculo de cotización (mapeado por fila real, igual que Egresos)
    if (iCotiz > -1 && rows.length) {
      try {
        var byRowC = {};
        for (var bc=0; bc<rows.length; bc++) byRowC[rows[bc].rowNum] = rows[bc];
        var rtC = sh.getRange(2, iCotiz+1, raw.length-1, 1).getRichTextValues();
        for (var rcc=0; rcc<rtC.length; rcc++){
          var uC = rtC[rcc][0] ? rtC[rcc][0].getLinkUrl() : '';
          if (uC && byRowC[rcc+2]) byRowC[rcc+2].linkCotizacionUrl = uC;
        }
      } catch(_eC){}
    }
    // Hipervínculo de factura (col S=19) y comprobante de pago (col T=20) —
    // uploadFile() los escribe como rich-text hyperlink, igual que cotización;
    // sin esta extracción el valor plano de la celda no sirve para abrir el visor.
    if (rows.length) {
      try {
        var byRowF = {};
        for (var bf=0; bf<rows.length; bf++) byRowF[rows[bf].rowNum] = rows[bf];
        var rtF = sh.getRange(2, 19, raw.length-1, 1).getRichTextValues();
        var rtP = sh.getRange(2, 20, raw.length-1, 1).getRichTextValues();
        for (var rcf=0; rcf<rtF.length; rcf++){
          var uF = rtF[rcf][0] ? rtF[rcf][0].getLinkUrl() : '';
          if (uF && byRowF[rcf+2]) byRowF[rcf+2].linkFacturaUrl = uF;
          var uP = rtP[rcf][0] ? rtP[rcf][0].getLinkUrl() : '';
          if (uP && byRowF[rcf+2]) byRowF[rcf+2].linkPagoUrl = uP;
        }
      } catch(_eF){}
    }
    var urgOrder = {vencido:0,hoy:1,semana:2,mes:3,ok:4,'sin-fecha':5};
    rows.sort(function(a,b){
      var ua=urgOrder[a.urgencia]||4, ub=urgOrder[b.urgencia]||4;
      if(ua!==ub) return ua-ub;
      return (a.dias||999)-(b.dias||999);
    });

    // Saldo pendiente real (Monto - abonos activos) y crédito disponible por
    // proveedor — para poder aplicarlo directo desde la lista sin adivinar.
    var creditosPorProveedor = {};
    try {
      var saldos = readSaldosCxP(rows.map(function(r){ return r.id; }));
      var abonadoPorId = (saldos && saldos.abonadoPorId) || {};
      rows.forEach(function(r){
        var abonado = abonadoPorId[r.id] || 0;
        r.montoAbonado = abonado;
        r.saldoPendiente = Math.max(0, r.monto - abonado);
      });
      var todosCreditos = readCreditosProveedor(); // sin filtro: un solo read para todos
      (todosCreditos && todosCreditos.rows || []).forEach(function(cr){
        creditosPorProveedor[cr.proveedor] = (creditosPorProveedor[cr.proveedor] || 0) + cr.montoDisponible;
      });
    } catch(eSaldo) {}

    return {ok:true, rows:rows, resumen:resumen, creditosPorProveedor: creditosPorProveedor};
  } catch(ex) { return {ok:false, error:ex.message, rows:[], resumen:{}}; }
}

function updateCxP(body) {
  // Edita directamente en Egresos2026 (misma fila que lee readBDCxP)
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var sh = ss.getSheetByName(egTab);
    if (!sh) return {ok:false, error:'Hoja Egresos no encontrada'};
    var rn = body.rowNum;
    if (!rn || rn < 2) return {ok:false, error:'Fila inválida'};
    var oldRow = sh.getRange(rn, 1, 1, 20).getValues()[0];
    var oldProv = String(oldRow[4]||'');
    // Cols Egresos: E=Proveedor(5), F=Contable(6), G=Tipo(7), H=Subtipo(8),
    // I=Concepto(9), J=Monto(10), K=Notas(11), L=Vencimiento(12),
    // P=Poliza(16), R=Obs(18)
    if (body.proveedor!==undefined) sh.getRange(rn,5).setValue(body.proveedor);
    if (body.contable!==undefined) sh.getRange(rn,6).setValue(body.contable);
    if (body.tipo!==undefined) sh.getRange(rn,7).setValue(body.tipo);
    if (body.subtipo!==undefined) sh.getRange(rn,8).setValue(body.subtipo);
    if (body.concepto!==undefined) sh.getRange(rn,9).setValue(body.concepto);
    if (body.monto!==undefined) sh.getRange(rn,10).setValue(parseFloat(String(body.monto||'').replace(/[$,]/g,''))||0);
    if (body.notas!==undefined) sh.getRange(rn,11).setValue(body.notas);
    if (body.vencimiento!==undefined) sh.getRange(rn,12).setValue(body.vencimiento);
    if (body.poliza!==undefined) sh.getRange(rn,16).setValue(body.poliza);
    if (body.observaciones!==undefined) sh.getRange(rn,18).setValue(body.observaciones);
    if (body.divisa!==undefined) { try { _cxpGuardarDivisa(sh, rn, body); } catch(_d) {} }
    logAudit(body.usuario||'sistema','CxP','Editar',String(oldRow[0]),'Proveedor',oldProv,body.proveedor||oldProv);
    return {ok:true, rowNum:rn};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function deleteCxPRow(body) {
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var sh = ss.getSheetByName(egTab);
    if (!sh) return {ok:false, error:'Hoja Egresos no encontrada'};
    var rn = parseInt(body.rowNum);
    if (!rn || rn < 2) return {ok:false, error:'Fila inválida'};
    var rowVals = sh.getRange(rn, 1, 1, 5).getValues()[0];
    var prov = String(rowVals[4]||'');
    sh.deleteRow(rn);
    logAudit(body.usuario||'sistema','CxP','Borrar',String(rowVals[0]||rn),'Proveedor',prov,'—');
    try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch(e) {}
    return {ok:true, rowNum:rn};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function pagarCxP(body) {
  // Actualiza la MISMA fila en Egresos2026: pone fecha + PAGADO + forma de pago
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var sh = ss.getSheetByName(egTab);
    if (!sh) return {ok:false, error:'Hoja Egresos no encontrada'};
    var rowNum = body.rowNum;
    if (!rowNum || rowNum < 2) return {ok:false, error:'Fila inválida'};

    var fechaPago = body.fechaPago || new Date().toISOString().substring(0,10);
    var formaPago = body.formaPago || '';

    // Detectar columnas por ENCABEZADO (robusto ante cambios de orden en la hoja).
    // Fallback a las posiciones históricas si no se encuentra el header. 0-based.
    var _hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var _col = function(sub, fb){ var lc = String(sub).toLowerCase(); for (var c = 0; c < _hdrs.length; c++) { if (_hdrs[c].indexOf(lc) > -1) return c; } return fb; };
    // FECHA DE PAGO: la hoja puede tener varias columnas con "fecha" (ej. Vencimiento).
    // 1) columna explícita de pago ('fecha' + 'pago'); 2) 'fecha' que NO sea vencimiento
    //    ni factura; 3) fallback col B. NUNCA escribe en la de vencimiento.
    var iFecha = (function(){
      for (var c=0;c<_hdrs.length;c++){ if (_hdrs[c].indexOf('fecha')>-1 && _hdrs[c].indexOf('pago')>-1) return c; }
      for (var c2=0;c2<_hdrs.length;c2++){ if (_hdrs[c2].indexOf('fecha')>-1 && _hdrs[c2].indexOf('venc')<0 && _hdrs[c2].indexOf('factur')<0) return c2; }
      return 1;
    })();
    var iPagado=_col('pagado',13), iForma=_col('forma',16),
        iEgresos=_col('egresos',9), iNotas=_col('notas',10), iMes=_col('mes',2),
        iProv=_col('proveedor',4), iConc=_col('concepto',8);

    var _fd = new Date(fechaPago);
    sh.getRange(rowNum, iFecha+1).setValue(_fd);          // Fecha de pago (por header)
    sh.getRange(rowNum, iPagado+1).setValue(true);        // PAGADO = TRUE
    if (formaPago) sh.getRange(rowNum, iForma+1).setValue(formaPago);
    // Mes tag consistente con la fecha (evita que quede desalineado con el filtro Mes)
    if (iMes > -1 && !isNaN(_fd)) {
      var _M3 = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      sh.getRange(rowNum, iMes+1).setValue(_M3[_fd.getMonth()] + '-' + String(_fd.getFullYear()).slice(-2));
    }
    // NOTA: el comprobante/factura y su hipervínculo los escribe uploadFile al subir el
    // archivo (único escritor del link). pagarCxP ya NO toca esas columnas, para no
    // sobrescribir con un link viejo (que el anti-duplicados pudo mandar a papelera).

    // Divisa (solo aplica a Santander / AMEX)
    var divisa = body.divisa || 'MXN';
    var tipoCambio = parseFloat(body.tipoCambio) || 0;
    var montoUSD = parseFloat(body.montoUSD) || 0;

    // Leer datos de la fila para banco routing (por header)
    var rowData = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
    var monto = parseFloat(String(rowData[iEgresos]||'').replace(/[$,\s]/g,'')) || 0;
    // Saldo parcial (créditos/abonos ya aplicados) — si viene, ese es el
    // importe real que se rutea a banco/caja; el Monto del egreso no se toca.
    if (body.montoOverride !== undefined && body.montoOverride !== null && parseFloat(body.montoOverride) > 0) {
      monto = parseFloat(body.montoOverride);
    }
    var proveedor = String(rowData[iProv] || '');
    var concepto = String(rowData[iConc] || '');
    var egId = String(rowData[0] || '');

    // Pago en USD: convertir a MXN, actualizar el monto del egreso y dejar nota
    var esUSD = (formaPago === 'Santander' || formaPago === 'AMEX') && divisa === 'USD' && tipoCambio > 0 && montoUSD > 0;
    if (esUSD) {
      monto = Math.round(montoUSD * tipoCambio * 100) / 100;
      sh.getRange(rowNum, iEgresos+1).setValue(monto);   // Monto del egreso en MXN
      var notaPrev = String(rowData[iNotas] || '');
      var notaUSD = 'USD ' + montoUSD.toFixed(2) + ' @ TC ' + tipoCambio;
      sh.getRange(rowNum, iNotas+1).setValue(notaPrev ? (notaPrev + ' · ' + notaUSD) : notaUSD);
    }

    // Rutear a banco según forma de pago
    var mesStr = fechaPago.substring(0, 7);
    var usdAmt = esUSD ? montoUSD : '';
    var tc     = esUSD ? tipoCambio : '';
    var ref    = concepto + ' · ' + proveedor + ' [Egreso #' + egId + ']';
    if (formaPago === 'Efectivo' && monto > 0) {
      // Efectivo → Caja Chica (usa getCajaChicaSheet para apuntar al SS correcto)
      try {
        var ccSh = getCajaChicaSheet();
        var ccData = ccSh.getDataRange().getValues();
        var ccHdr  = ccData[0].map(function(h){ return String(h).trim().toUpperCase(); });
        var ciF = ccHdr.indexOf('FECHA'), ciC = ccHdr.indexOf('CONCEPTO'), ciS = ccHdr.indexOf('SALIDA');
        // Buscar primera fila vacía o append
        var ccRow = -1;
        for (var ci = 1; ci < ccData.length; ci++) {
          if (!String(ccData[ci][ciF]||'').trim() && !String(ccData[ci][ciC]||'').trim()) { ccRow = ci + 1; break; }
        }
        if (ccRow === -1) ccRow = ccSh.getLastRow() + 1;
        ccSh.getRange(ccRow, ciF+1).setValue(fechaPago);
        ccSh.getRange(ccRow, ciC+1).setValue(ref);
        ccSh.getRange(ccRow, ciS+1).setValue(monto);
        SpreadsheetApp.flush();
      } catch(ccErr) { Logger.log('pagarCxP Efectivo→CajaChica error: ' + ccErr.message); }
    } else if (formaPago && monto > 0) {
      var banco = '', bankRow = null;
      if (formaPago === 'Santander' || formaPago === 'Transferencia') {
        banco = 'santander';
        bankRow = [fechaPago, 0, monto, 0, ref, usdAmt, tc, '', ''];
      } else if (formaPago === 'AMEX') {
        banco = 'amex';
        bankRow = [fechaPago, monto, 0, ref, usdAmt, tc, '', '', mesStr];
      } else if (formaPago === 'Mercado Pago' || formaPago === 'TDC' || formaPago === 'TDD') {
        banco = 'mercadopago';
        // Egreso por MP = RETIRO. Cobro y Neto negativos para que la conciliación
        // lo muestre COMPLETO (badge Retiro + importe visible), no solo en Neto.
        bankRow = [mesStr, fechaPago, -monto, 0, 0, -monto, 0, false, ref, 'PAGO'];
      }
      if (banco && bankRow) {
        try { saveBankRow(banco, bankRow); } catch(bErr) { Logger.log('pagarCxP banco error: ' + bErr.message); }
      }
    }

    logAudit(body.usuario || 'sistema', 'CxP', 'Pagar', egId, 'Pagado', 'Pendiente', formaPago + ' | ' + fechaPago);
    return {ok: true, egresoId: egId, rowNum: rowNum,
      _debug: { colFecha: (iFecha+1), headerFecha: (_hdrs[iFecha]||''), fechaPagoEscrita: fechaPago } };
  } catch (ex) { return {ok: false, error: ex.message}; }
}

function migrateCxPFromEgresos() {
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var egSh = ss.getSheetByName(egTab);
    if (!egSh) return {ok:false, error:'Hoja Egresos no encontrada'};
    setupBDCxP();
    var cxpSh = ss.getSheetByName(BD_CXP_TAB);

    var raw = egSh.getDataRange().getValues();
    function dt(v){if(!v)return'';if(v instanceof Date)return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');return String(v);}
    function num(v){var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}

    var counter = 0;
    var lr = cxpSh.getLastRow();
    if (lr > 1) {
      var last = String(cxpSh.getRange(lr,1).getValue()||'');
      var m = last.match(/CXP-(\d+)/);
      if (m) counter = parseInt(m[1],10);
    }

    var batchRows = [];
    for (var i=1;i<raw.length;i++) {
      var r = raw[i];
      var fecha = r[1]; // col B = Fecha de pago
      var pagado = r[13]===true||String(r[13]).toUpperCase()==='TRUE';
      // Si no tiene fecha Y no está pagado → es CxP
      if (!fecha && !pagado && String(r[0]||'').trim()) {
        counter++;
        var cxpId = 'CXP-'+String(counter).padStart(5,'0');
        batchRows.push([
          cxpId, String(r[2]||''), num(r[3]),
          String(r[4]||''), String(r[5]||''), String(r[6]||''), String(r[7]||''),
          String(r[8]||''), num(r[9]), String(r[10]||''),
          dt(r[11]), // Vencimiento
          r[12]===true||String(r[12]).toUpperCase()==='TRUE', // Facturacion
          false, // Pagado
          r[14]===true||String(r[14]).toUpperCase()==='TRUE', // Contabilidad
          String(r[15]||''), String(r[16]||''), String(r[17]||''),
          String(r[18]||''), String(r[19]||''),
          new Date()
        ]);
      }
    }
    if (batchRows.length) {
      cxpSh.getRange(cxpSh.getLastRow()+1, 1, batchRows.length, batchRows[0].length).setValues(batchRows);
    }
    return {ok:true, migrated:batchRows.length, lastId:'CXP-'+String(counter).padStart(5,'0')};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

/* ══════════════════════════════════════════════════════════════
   INGRESOS — lectura de spreadsheet mensual de ingresos
   ══════════════════════════════════════════════════════════════ */
var INGRESOS_SS_ID = '1x_TE_YxLOwnBXKV_lA3Ss_EOSu1p61uTdUmw2zh_6uc'; // Ingresos 2026
var INGRESOS_SS_2025 = '17gNzXavMbQ8DhFEIxCqzCJ6z-wTZgIwL4ibEDyVKE2w';
var INGRESOS_SS_2024 = '1Zx4QWulAgrrVBeI8nfTR10EiYL-l3crJsaZSYDbWSug';
var PRODUCTOS_SS_ID = '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY';
var PACIENTES_SS_ID = '1uoQU-vbefxWwaLxJyTFT25gj7Nr2223WISa3tqH-Rio';

/* ══════════════════════════════════════════════════════════════
   EGRESOS — lectura y captura de gastos/costos
   ══════════════════════════════════════════════════════════════ */
var EGRESOS_SS_2026 = '1iRjpYtkcqx-3NRwlVK-UYx09I0gVyiTDRtIA9X9RAQw';
var EGRESOS_SS_2025 = '18Wf4tD6CYBMTGVLPkEw_5YOtJncOAAeCyfeKMe--M1g';
var EGRESOS_SS_2024 = '18DOfh1CvMyY3ZntjXGEqw6mjzhhBYZkatygxvnck2Is';
var EGRESOS_TABS = { 2026:'Egresos2026', 2025:'Egresos2025', 2024:'Egresos2024' };
var EGRESOS_IDS  = { 2026:EGRESOS_SS_2026, 2025:EGRESOS_SS_2025, 2024:EGRESOS_SS_2024 };
// Carpetas raíz en G:\...\01 Administración y Finanzas\01 Contabilidad
// Dentro de cada una se crean subcarpetas <Año>\<Mes> automáticamente.
var EGRESOS_DRIVE_FACTURAS = '1t8--HM1xymgqGyBbIsI2jhMVCgQUBm9n'; // Contabilidad\Facturas Recibidas
var EGRESOS_DRIVE_PAGOS    = '1D9H3nNIrkgg2wqJtKXzhuSLDH6hIUoPk'; // Contabilidad\Pagos
var EGRESOS_DRIVE_COTIZACIONES = '1o8J61IsrlaTBoENtwQlIA_ADTZeOLas1'; // Contabilidad\Cotizaciones

// Devuelve la columna (1-indexed) cuyo header contiene `want`; si no existe, la crea al final.
function _egColEnsure(sh, want, headerText) {
  var lastCol = sh.getLastColumn();
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); });
  for (var c = 0; c < hdrs.length; c++) { if (hdrs[c].indexOf(want) > -1) return c + 1; }
  sh.getRange(1, lastCol + 1).setValue(headerText);
  return lastCol + 1;
}
var EGRESOS_MESES_FOLDER = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function _getOrCreateMonthFolder(parentFolderId, anio, mes) {
  var parent = DriveApp.getFolderById(parentFolderId);
  var mesName = EGRESOS_MESES_FOLDER[mes] || 'Mes' + (mes + 1);
  // Buscar o crear carpeta del año
  var yearName = String(anio);
  var yearFolder = null;
  var yearIter = parent.getFoldersByName(yearName);
  if (yearIter.hasNext()) { yearFolder = yearIter.next(); }
  else { yearFolder = parent.createFolder(yearName); }
  // Buscar o crear carpeta del mes
  var mesFolder = null;
  var mesIter = yearFolder.getFoldersByName(mesName);
  if (mesIter.hasNext()) { mesFolder = mesIter.next(); }
  else { mesFolder = yearFolder.createFolder(mesName); }
  return mesFolder;
}

function uploadEgresoPDF(payload) {
  try {
    var tipo = payload.tipo || 'factura'; // 'factura' o 'pago'
    var parentId = tipo === 'pago' ? EGRESOS_DRIVE_PAGOS : EGRESOS_DRIVE_FACTURAS;
    var base64 = payload.base64;
    var fileName = payload.fileName || 'documento.pdf';
    var rowNum = parseInt(payload.rowNum);
    var anio = payload.anio || new Date().getFullYear();

    if (!base64) return {ok:false, error:'No se recibió archivo'};
    if (!rowNum || rowNum < 2) return {ok:false, error:'Fila inválida'};

    // Determinar mes desde la fecha del egreso
    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = EGRESOS_TABS[anio] || 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === tabName) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'Pestaña no encontrada'};

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){return String(h).trim().toLowerCase();});
    var iFecha = -1;
    for (var c = 0; c < headers.length; c++) { if (headers[c].indexOf('fecha') > -1) { iFecha = c; break; } }

    var fechaVal = iFecha > -1 ? sheet.getRange(rowNum, iFecha + 1).getValue() : new Date();
    var fechaObj = fechaVal instanceof Date ? fechaVal : new Date(fechaVal);
    if (isNaN(fechaObj.getTime())) fechaObj = new Date();

    var mes = fechaObj.getMonth();
    var anioFile = fechaObj.getFullYear();

    // Crear/obtener carpeta mes
    var folder = _getOrCreateMonthFolder(parentId, anioFile, mes);

    // Subir archivo
    var decoded = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(decoded, 'application/pdf', fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();

    // Guardar URL como hipervínculo en la celda correspondiente
    var iLinkCol = -1;
    var linkHeader = tipo === 'pago' ? 'link pago' : 'link factura';
    for (var lc = 0; lc < headers.length; lc++) {
      if (headers[lc].indexOf(linkHeader) > -1) { iLinkCol = lc; break; }
    }

    var displayName = fileName.replace(/\.pdf$/i, '');
    if (iLinkCol > -1) {
      var richText = SpreadsheetApp.newRichTextValue()
        .setText(displayName)
        .setLinkUrl(fileUrl)
        .build();
      sheet.getRange(rowNum, iLinkCol + 1).setRichTextValue(richText);
    }

    // Auto-activar checkbox correspondiente
    var checkHeader = tipo === 'pago' ? 'pagado' : 'facturación';
    // Buscar con y sin tilde
    var iCheck = -1;
    for (var ch = 0; ch < headers.length; ch++) {
      if (headers[ch].replace(/[áàä]/g,'a').replace(/[óòö]/g,'o').indexOf(checkHeader.replace(/[áàä]/g,'a').replace(/[óòö]/g,'o')) > -1) {
        iCheck = ch; break;
      }
    }
    if (iCheck > -1) {
      sheet.getRange(rowNum, iCheck + 1).setValue(true);
    }

    return {ok:true, url:fileUrl, fileName:displayName, tipo:tipo, rowNum:rowNum};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function readEgresosData(anio) {
  try {
    anio = anio || new Date().getFullYear();
    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = EGRESOS_TABS[anio] || 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === tabName) { sheet = sheets[i]; break; }
    }
    if (!sheet) sheet = sheets[0];

    var raw = sheet.getDataRange().getValues();
    if (raw.length < 2) return {ok:true, view:'egresos', rows:[], totalRows:0, proveedores:[], subtipos:[], anio:anio};

    // Detectar headers (fila 0)
    var headers = raw[0].map(function(h){ return String(h).trim(); });
    function col(name) {
      var lc = name.toLowerCase();
      for (var c = 0; c < headers.length; c++) { if (headers[c].toLowerCase().indexOf(lc) > -1) return c; }
      return -1;
    }
    var iFecha=col('fecha'), iMes=col('mes'), iProveedor=col('proveedor'), iContable=col('contable'),
        iTipo=col('tipo'), iSubtipo=col('subtipo'), iConcepto=col('concepto'), iEgresos=col('egresos'),
        iNotas=col('notas'), iVenc=col('vencimiento'), iFact=col('facturación'), iPagado=col('pagado'),
        iCont=col('contabilidad'), iPoliza=col('póliza')===-1?col('poliza'):col('póliza'),
        iFPago=col('forma de pago'), iObs=col('observaciones'), iLinkFact=col('link factura'),
        iLinkPago=col('link pago'), iLinkCotiz=col('cotiz'),
        iUSD=col('usd'), iTC=col('tipo de cambio'), iEstatus=col('estatus'),
        iRec=col('recurrente'), iDeveng=col('devengado');

    function num(v) { if (typeof v==='number') return v; var n=parseFloat(String(v||'').replace(/[$,\s]/g,'')); return isNaN(n)?0:n; }
    function dt(v) { if(!v)return''; if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0'); return String(v); }
    function bool(v) { return v===true||String(v).toUpperCase()==='TRUE'; }

    var allRows = [];
    var provSet = {}, subtipoSet = {}, contableSet = {}, fpSet = {};

    for (var r = 1; r < raw.length; r++) {
      var row = raw[r];
      var proveedor = iProveedor>-1 ? String(row[iProveedor]||'').trim() : '';
      var concepto  = iConcepto>-1 ? String(row[iConcepto]||'').trim() : '';
      if (!proveedor && !concepto) continue; // skip empty/summary rows
      var monto = iEgresos>-1 ? num(row[iEgresos]) : 0;

      var fecha = iFecha>-1 ? dt(row[iFecha]) : '';
      var subtipo = iSubtipo>-1 ? String(row[iSubtipo]||'').trim() : '';
      var contable = iContable>-1 ? String(row[iContable]||'').trim() : '';
      var tipo = iTipo>-1 ? String(row[iTipo]||'').trim() : '';
      var fp = iFPago>-1 ? String(row[iFPago]||'').trim() : '';

      if (proveedor) provSet[proveedor] = 1;
      if (subtipo) subtipoSet[subtipo] = 1;
      if (contable) contableSet[contable] = 1;
      if (fp) fpSet[fp] = 1;

      allRows.push({
        _rowNum: r + 1,
        id: String(row[0]||'').trim(),
        fecha: fecha,
        mes: iMes>-1 ? String(row[iMes]||'').trim() : '',
        proveedor: proveedor,
        contable: contable,
        tipo: tipo,
        subtipo: subtipo,
        concepto: concepto,
        monto: monto,
        notas: iNotas>-1 ? String(row[iNotas]||'').trim() : '',
        vencimiento: iVenc>-1 ? dt(row[iVenc]) : '',
        facturacion: iFact>-1 ? bool(row[iFact]) : false,
        pagado: iPagado>-1 ? bool(row[iPagado]) : false,
        contabilidad: iCont>-1 ? bool(row[iCont]) : false,
        poliza: iPoliza>-1 ? String(row[iPoliza]||'').trim() : '',
        formaPago: fp,
        observaciones: iObs>-1 ? String(row[iObs]||'').trim() : '',
        linkFactura: iLinkFact>-1 ? String(row[iLinkFact]||'').trim() : '',
        linkFacturaUrl: '',
        linkPago: iLinkPago>-1 ? String(row[iLinkPago]||'').trim() : '',
        linkPagoUrl: '',
        linkCotizacion: iLinkCotiz>-1 ? String(row[iLinkCotiz]||'').trim() : '',
        linkCotizacionUrl: '',
        montoUSD: iUSD>-1 ? num(row[iUSD]) : 0,
        tipoCambio: iTC>-1 ? num(row[iTC]) : 0,
        estatus: iEstatus>-1 ? String(row[iEstatus]||'').trim() : '',
        recurrenteId: iRec>-1 ? String(row[iRec]||'').trim() : '',
        mesDevengado: iDeveng>-1 ? (function(v){ if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0'); return String(v||'').trim(); })(row[iDeveng]) : ''
      });
    }

    // Extraer URLs de hipervínculos (Link Factura / Link Pago).
    // IMPORTANTE: el bucle de arriba se salta filas vacías, así que NO se puede mapear
    // por posición de array. Mapeamos por fila REAL de la hoja usando _rowNum.
    try {
      var lastDataRow = raw.length - 1;
      if (lastDataRow > 0 && (iLinkFact > -1 || iLinkPago > -1 || iLinkCotiz > -1)) {
        var byRow = {};
        for (var ar = 0; ar < allRows.length; ar++) byRow[allRows[ar]._rowNum] = allRows[ar];
        if (iLinkCotiz > -1) {
          var rtCot = sheet.getRange(2, iLinkCotiz + 1, lastDataRow, 1).getRichTextValues();
          for (var rc = 0; rc < rtCot.length; rc++) {
            var urlC = rtCot[rc][0] ? rtCot[rc][0].getLinkUrl() : '';
            if (urlC && byRow[rc + 2]) byRow[rc + 2].linkCotizacionUrl = urlC; // rc+2 = fila real
          }
        }
        if (iLinkFact > -1) {
          var rtFact = sheet.getRange(2, iLinkFact + 1, lastDataRow, 1).getRichTextValues();
          for (var rf = 0; rf < rtFact.length; rf++) {
            var url = rtFact[rf][0] ? rtFact[rf][0].getLinkUrl() : '';
            if (url && byRow[rf + 2]) byRow[rf + 2].linkFacturaUrl = url; // rf+2 = fila real
          }
        }
        if (iLinkPago > -1) {
          var rtPago = sheet.getRange(2, iLinkPago + 1, lastDataRow, 1).getRichTextValues();
          for (var rp = 0; rp < rtPago.length; rp++) {
            var url2 = rtPago[rp][0] ? rtPago[rp][0].getLinkUrl() : '';
            if (url2 && byRow[rp + 2]) byRow[rp + 2].linkPagoUrl = url2; // rp+2 = fila real
          }
        }
      }
    } catch(rtErr) { /* getRichTextValues puede fallar en sheets muy grandes — continuamos sin URLs */ }

    // KPIs
    var totalEgresos=0, totalPagado=0, totalPendiente=0, countPagado=0;
    var subtipoMap={}, mesMap={}, provMap={}, contMap={};
    allRows.forEach(function(r) {
      totalEgresos += r.monto;
      if (r.pagado) { totalPagado += r.monto; countPagado++; }
      else { totalPendiente += r.monto; }
      var st = r.subtipo||'Sin subtipo';
      if(!subtipoMap[st]) subtipoMap[st]={nombre:st,total:0,count:0}; subtipoMap[st].total+=r.monto; subtipoMap[st].count++;
      var m = r.mes||'Sin mes';
      if(!mesMap[m]) mesMap[m]={mes:m,total:0,count:0}; mesMap[m].total+=r.monto; mesMap[m].count++;
      var p = r.proveedor||'Sin proveedor';
      if(!provMap[p]) provMap[p]={nombre:p,total:0,count:0}; provMap[p].total+=r.monto; provMap[p].count++;
      var c = r.contable||'Sin clasificar';
      if(!contMap[c]) contMap[c]={nombre:c,total:0}; contMap[c].total+=r.monto;
    });

    function topArr(map,limit) {
      var arr=[]; for(var k in map) arr.push(map[k]);
      arr.sort(function(a,b){return b.total-a.total;});
      arr.forEach(function(i){i.pct=totalEgresos>0?(i.total/totalEgresos)*100:0;});
      return arr.slice(0,limit||10);
    }

    return {
      ok:true, view:'egresos', anio:anio,
      rows: allRows.reverse(), // más reciente primero
      totalRows: allRows.length,
      totalEgresos: totalEgresos,
      totalPagado: totalPagado,
      totalPendiente: totalPendiente,
      promedioMensual: Object.keys(mesMap).length>0 ? totalEgresos/Object.keys(mesMap).length : 0,
      topSubtipos: topArr(subtipoMap, 15),
      topProveedores: topArr(provMap, 15),
      distribucionContable: topArr(contMap, 5),
      resumenMensual: topArr(mesMap, 12),
      proveedores: Object.keys(provSet).sort(),
      subtipos: Object.keys(subtipoSet).sort(),
      contables: Object.keys(contableSet).sort(),
      formasPago: Object.keys(fpSet).sort()
    };
  } catch(ex) {
    return {ok:false, error:ex.message, rows:[], totalRows:0};
  }
}

/* ══════════════════════════════════════════════════════════════
   EGRESOS — Setup de hoja por año y reporte multi-año
   ══════════════════════════════════════════════════════════════ */
function setupEgresosAnio(anio) {
  anio = parseInt(anio) || new Date().getFullYear();
  var ssId = EGRESOS_IDS[anio];
  if (!ssId) return {ok:false, error:'Año no configurado: '+anio};
  var tabName = EGRESOS_TABS[anio] || 'Egresos'+anio;
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var sh = ss.getSheetByName(tabName);
    var created = false;
    if (!sh) {
      sh = ss.insertSheet(tabName);
      created = true;
    }
    var raw = sh.getDataRange().getValues();
    var rows = Math.max(0, raw.length - 1);
    var hasHeaders = raw.length > 0 && raw[0].filter(function(h){ return String(h).trim(); }).length > 3;
    if (!hasHeaders) {
      var hdrs = ['ID','Fecha','Mes','Proveedor','Contable','Tipo','Subtipo','Concepto',
        'Egresos','Notas','Vencimiento','Facturación','Pagado','Contabilidad','Póliza',
        'Forma de Pago','Observaciones','Link Factura','Link Pago','Cotización',
        'Monto USD','Tipo de Cambio'];
      sh.getRange(1,1,1,hdrs.length).setValues([hdrs]);
      sh.getRange(1,1,1,hdrs.length).setFontWeight('bold').setBackground('#fce7f3');
      sh.setFrozenRows(1);
      created = true; rows = 0;
    } else {
      // Migración segura: agrega columnas faltantes AL FINAL, sin mover nada existente
      var existing = raw[0].map(function(h){ return String(h).trim().toLowerCase(); });
      var toAdd = [];
      if (existing.indexOf('monto usd') === -1) toAdd.push('Monto USD');
      if (existing.indexOf('tipo de cambio') === -1) toAdd.push('Tipo de Cambio');
      if (toAdd.length) {
        var lastCol = sh.getLastColumn();
        sh.getRange(1, lastCol+1, 1, toAdd.length).setValues([toAdd]);
        sh.getRange(1, lastCol+1, 1, toAdd.length).setFontWeight('bold').setBackground('#fce7f3');
      }
    }
    return {ok:true, anio:anio, tabName:tabName, ssId:ssId, created:created, rows:rows};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function readReporteProveedor(body) {
  var prov = String(body.proveedor||'').trim();
  var fp   = String(body.formaPago||'').trim();
  var cont = String(body.contable||'').trim();
  var sub  = String(body.subtipo||'').trim();
  var ini  = String(body.ini||'');
  var fin  = String(body.fin||'');
  if (!prov && !fp && !cont && !sub) return {ok:false, error:'Selecciona al menos un filtro (proveedor, forma de pago, contable o subtipo).'};
  // Normaliza para comparar proveedores entre años (quita acentos, puntos, dobles espacios)
  function _normProv(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,' ').replace(/[.,]/g,'').replace(/[áàä]/g,'a').replace(/[éèë]/g,'e').replace(/[íìï]/g,'i').replace(/[óòö]/g,'o').replace(/[úùü]/g,'u'); }
  var provNorm = _normProv(prov);
  var provPfx = provNorm.substring(0, 5);   // prefijo para muestrear variantes del nombre
  // Coincide si es igual, o si uno contiene al otro (tolera sufijos/prefijos entre años)
  function _provMatch(rprov){
    if (!prov) return true;
    var rn = _normProv(rprov); if (!rn) return false;
    return rn === provNorm || rn.indexOf(provNorm) > -1 || provNorm.indexOf(rn) > -1;
  }
  var curYear = new Date().getFullYear();
  var allRows = [];
  var diag = [];   // diagnóstico por año (para ver por qué no encuentra)
  var years = Object.keys(EGRESOS_IDS).map(Number).sort();
  years.forEach(function(anio) {
    var dg = { anio:anio, ok:false, totalRows:0, matched:0, variantes:[] };
    try {
      var eg = readEgresosData(anio);
      if (!eg.ok || !eg.rows) { dg.error = (eg && eg.error) || 'sin filas'; diag.push(dg); return; }
      dg.ok = true; dg.totalRows = eg.rows.length;
      var esHistoricoAnio = anio < curYear;
      eg.rows.forEach(function(r) {
        // Muestrea proveedores parecidos (mismo prefijo) para detectar variantes de nombre
        if (provPfx && _normProv(r.proveedor).indexOf(provPfx) > -1 && dg.variantes.indexOf(r.proveedor) < 0 && dg.variantes.length < 6) dg.variantes.push(r.proveedor);
        // Cuenta como gasto realizado si: pagado, o con fecha de pago, o —para años
        // pasados (2024/2025, formato viejo sin casilla "pagado")— cualquier monto.
        if (!r.pagado && !(r.fecha||'') && !(esHistoricoAnio && (r.monto||0)>0)) return;
        if (r.estatus === 'Cancelada') return;
        if (!_provMatch(r.proveedor)) return;
        if (fp && (r.formaPago||'').trim() !== fp) return;
        if (cont && (r.contable||'').trim() !== cont) return;
        if (sub && (r.subtipo||'').trim() !== sub) return;
        var fd = (r.fecha||'').substring(0,10);
        if (ini && fd && fd < ini) return;
        if (fin && fd && fd > fin) return;
        dg.matched++; allRows.push(r);
      });
    } catch(e) { dg.error = e.message; }
    diag.push(dg);
  });
  if (!allRows.length) return {ok:true, proveedor:prov, rows:[], diag:diag,
    totales:{total:0,count:0,avg:0,ultimoPago:'',ultimaFormaPago:''},tendencia:[]};
  allRows.sort(function(a,b){ return (a.fecha||'').localeCompare(b.fecha||''); });
  var total = allRows.reduce(function(s,r){ return s+(r.monto||0); }, 0);
  var count = allRows.length;
  var last = allRows[allRows.length-1];
  var byMes = {};
  allRows.forEach(function(r){
    var k=(r.fecha||'').substring(0,7); if(k) byMes[k]=(byMes[k]||0)+(r.monto||0);
  });
  var MN=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var tendencia = Object.keys(byMes).sort().map(function(k){
    var p=k.split('-'); return {mes:k, label:(MN[parseInt(p[1],10)-1]||p[1])+' '+(p[0]||'').substring(2), monto:byMes[k]};
  });
  return {ok:true, proveedor:prov, ini:ini, fin:fin, rows:allRows, diag:diag,
    totales:{total:total, count:count, avg:count?total/count:0,
      ultimoPago:last.fecha||'', ultimaFormaPago:last.formaPago||''},
    tendencia:tendencia};
}

function saveEgreso(payload) {
  try {
    var anio = payload.anio || new Date().getFullYear();
    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = EGRESOS_TABS[anio] || 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === tabName) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'Pestaña ' + tabName + ' no encontrada'};

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
    function col(name) {
      var lc = name.toLowerCase();
      for (var c=0;c<headers.length;c++) { if(headers[c].toLowerCase().indexOf(lc)>-1) return c; }
      return -1;
    }

    // Encontrar última fila con datos reales (no resúmenes)
    var lastDataRow = sheet.getLastRow();

    // Construir nueva fila basada en headers
    var newRow = [];
    for (var h = 0; h < headers.length; h++) newRow.push('');

    var iFecha=col('fecha'), iMes=col('mes'), iProveedor=col('proveedor'),
        iContable=col('contable'), iTipo=col('tipo'), iSubtipo=col('subtipo'),
        iConcepto=col('concepto'), iEgresos=col('egresos'), iNotas=col('notas'),
        iVenc=col('vencimiento'), iFact=col('facturación')===-1?col('facturacion'):col('facturación'),
        iPagado=col('pagado'), iCont=col('contabilidad'),
        iPoliza=col('póliza')===-1?col('poliza'):col('póliza'),
        iFPago=col('forma de pago'), iObs=col('observaciones'),
        iLinkFact=col('link factura'), iLinkPago=col('link pago'),
        iUSD=col('usd'), iTC=col('tipo de cambio');

    // Número de fila (#) en columna 0
    var iNum = -1;
    if (headers[0] && /^\d|#/i.test(headers[0])) iNum = 0;

    if (iNum>-1) newRow[iNum] = lastDataRow; // auto-número
    if (iFecha>-1) newRow[iFecha] = payload.fecha || '';
    if (iMes>-1) {
      var fd = new Date(payload.fecha);
      var meses = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      newRow[iMes] = isNaN(fd)?'':(meses[fd.getMonth()]+'-'+String(fd.getFullYear()).slice(-2));
    }
    if (iProveedor>-1) newRow[iProveedor] = payload.proveedor || '';
    if (iContable>-1) newRow[iContable] = payload.contable || '';
    if (iTipo>-1) newRow[iTipo] = payload.tipo || '';
    if (iSubtipo>-1) newRow[iSubtipo] = payload.subtipo || '';
    if (iConcepto>-1) newRow[iConcepto] = payload.concepto || '';
    if (iEgresos>-1) newRow[iEgresos] = parseFloat(String(payload.monto||'').replace(/[$,]/g,'')) || 0;
    if (iNotas>-1) newRow[iNotas] = payload.notas || '';
    if (iVenc>-1) newRow[iVenc] = payload.vencimiento || '';
    if (iFact>-1) newRow[iFact] = payload.facturacion === true || payload.facturacion === 'true';
    if (iPagado>-1) newRow[iPagado] = payload.pagado === true || payload.pagado === 'true';
    if (iCont>-1) newRow[iCont] = payload.contabilidad === true || payload.contabilidad === 'true';
    if (iPoliza>-1) newRow[iPoliza] = payload.poliza || '';
    if (iFPago>-1) newRow[iFPago] = payload.formaPago || '';
    if (iObs>-1) newRow[iObs] = payload.observaciones || '';
    if (iLinkFact>-1) newRow[iLinkFact] = payload.linkFactura || '';
    if (iLinkPago>-1) newRow[iLinkPago] = payload.linkPago || '';
    if (iUSD>-1) newRow[iUSD] = parseFloat(String(payload.montoUSD||'').replace(/[$,]/g,'')) || 0;
    if (iTC>-1) newRow[iTC] = parseFloat(String(payload.tipoCambio||'').replace(/[$,]/g,'')) || 0;

    // Column 1 (status) = 2 (capturado)
    if (headers.length > 3) {
      var iCol1 = -1;
      for (var ci=0;ci<headers.length;ci++) { if(/column|col\s*1|status/i.test(headers[ci])){iCol1=ci;break;} }
      if (iCol1 > -1) newRow[iCol1] = 2;
    }

    sheet.appendRow(newRow);
    var insertedRow = sheet.getLastRow();

    try { CacheService.getScriptCache().remove('gas_egresos_v1_' + anio); } catch(e) {}
    return {ok:true, rowNum:insertedRow, monto:newRow[iEgresos]};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function updateEgresoField(payload) {
  try {
    var anio = payload.anio || new Date().getFullYear();
    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = EGRESOS_TABS[anio] || 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === tabName) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'Pestaña no encontrada'};

    var rowNum = parseInt(payload.rowNum);
    if (!rowNum || rowNum < 2) return {ok:false, error:'Fila inválida'};

    var headersRaw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var headers = headersRaw.map(function(h){return String(h).trim().toLowerCase();});

    function findCol(name) {
      var lc = name.toLowerCase().replace(/[áàä]/g,'a').replace(/[éèë]/g,'e').replace(/[íìï]/g,'i').replace(/[óòö]/g,'o').replace(/[úùü]/g,'u');
      for (var c=0;c<headers.length;c++) {
        var hc = headers[c].replace(/[áàä]/g,'a').replace(/[éèë]/g,'e').replace(/[íìï]/g,'i').replace(/[óòö]/g,'o').replace(/[úùü]/g,'u');
        if (hc.indexOf(lc)>-1) return c;
      }
      return -1;
    }

    // Modo 1: edición completa desde formulario (tiene payload.fecha, payload.proveedor, etc.)
    if (payload.fecha && payload.proveedor) {
      var colMap = {
        'fecha':payload.fecha, 'proveedor':payload.proveedor, 'contable':payload.contable,
        'tipo':payload.tipo, 'subtipo':payload.subtipo, 'concepto':payload.concepto,
        'egresos':parseFloat(String(payload.monto||'').replace(/[$,]/g,''))||0,
        'notas':payload.notas||'', 'vencimiento':payload.vencimiento||'',
        'facturacion':payload.facturacion===true||payload.facturacion==='true',
        'pagado':payload.pagado===true||payload.pagado==='true',
        'contabilidad':payload.contabilidad===true||payload.contabilidad==='true',
        'poliza':payload.poliza||'', 'forma de pago':payload.formaPago||'',
        'observaciones':payload.observaciones||'',
        'usd':parseFloat(String(payload.montoUSD||'').replace(/[$,]/g,''))||0,
        'tipo de cambio':parseFloat(String(payload.tipoCambio||'').replace(/[$,]/g,''))||0
      };
      // Actualizar mes automáticamente
      var fd = new Date(payload.fecha);
      var meses = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var iMes = findCol('mes');
      if (iMes>-1 && !isNaN(fd)) sheet.getRange(rowNum, iMes+1).setValue(meses[fd.getMonth()]+'-'+String(fd.getFullYear()).slice(-2));

      for (var field in colMap) {
        var ci = findCol(field);
        if (ci > -1) sheet.getRange(rowNum, ci+1).setValue(colMap[field]);
      }

      try {
        logAudit(payload.usuario||'sistema', 'Egresos', 'Editar completo', 'Fila '+rowNum,
          payload.proveedor+' · '+payload.concepto, '', '$'+(colMap.egresos||0));
      } catch(ae){}

      return {ok:true, rowNum:rowNum, edited:true};
    }

    // Modo 2: campos individuales (checkboxes inline)
    var fields = payload.fields || {};
    for (var key in fields) {
      var ci2 = findCol(key);
      if (ci2 > -1) sheet.getRange(rowNum, ci2+1).setValue(fields[key]);
    }
    // Al marcar PAGADO por checkbox y no haber fecha de pago, poner la de hoy
    // (col Fecha) + actualizar Mes — para que no quede el egreso con fecha vacía.
    if (('pagado' in fields) && (fields.pagado === true || fields.pagado === 'true')) {
      var iFP = findCol('fecha');
      if (iFP > -1) {
        var actualF = sheet.getRange(rowNum, iFP + 1).getValue();
        if (!actualF || String(actualF).trim() === '') {
          var hoyP = new Date();
          sheet.getRange(rowNum, iFP + 1).setValue(hoyP);
          var mesesP = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
          var iMesP = findCol('mes');
          if (iMesP > -1) sheet.getRange(rowNum, iMesP + 1).setValue(mesesP[hoyP.getMonth()] + '-' + String(hoyP.getFullYear()).slice(-2));
        }
      }
    }

    return {ok:true, rowNum:rowNum};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

/* Repara egresos ya PAGADOS que quedaron SIN fecha (ej. se marcó la casilla PAG
   antes del arreglo, o el pago no pasó por pagarCxP). Rellena col Fecha con la
   fecha más sensata SIN adivinar mal el periodo:
     1) Vencimiento (si existe)
     2) el Mes registrado (col Mes, ej. "Jul-26") → día 15 de ese mes
     3) hoy (último recurso)
   No toca filas canceladas ni no pagadas. Devuelve cuántas reparó. */
function repararEgresosSinFecha(body) {
  try {
    var anio = (body && body.anio) || new Date().getFullYear();
    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = EGRESOS_TABS[anio] || 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(tabName) || ss.getSheets()[0];
    if (!sheet) return { ok:false, error:'Pestaña no encontrada' };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok:true, reparados:0, filas:[] };
    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    function col(sub){ for (var c=0;c<headers.length;c++){ if (headers[c].indexOf(sub)>-1) return c; } return -1; }
    var iFecha=col('fecha'), iMes=col('mes'), iVenc=col('vencimiento'), iPag=col('pagado'),
        iEstatus=col('estatus'), iProv=col('proveedor'), iMonto=col('egresos');
    if (iFecha < 0) return { ok:false, error:'No hay columna Fecha' };
    var MES3 = { ene:0,jan:0,feb:1,mar:2,abr:3,apr:3,may:4,jun:5,jul:6,ago:7,aug:7,sep:8,oct:9,nov:10,dic:11,dec:11 };
    var mesesOut = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var hoy = new Date();
    function parseMesTag(s){
      var m = String(s||'').trim().toLowerCase().match(/^([a-záéíóú]{3})[\s\-\/]*(\d{2,4})$/);
      if (!m) return null;
      var mo = MES3[m[1].replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i').replace(/ó/g,'o').replace(/ú/g,'u')];
      if (mo == null) return null;
      var y = m[2].length === 2 ? 2000 + parseInt(m[2],10) : parseInt(m[2],10);
      return new Date(y, mo, 15);
    }
    var reparados = 0, filas = [];
    for (var r=1; r<data.length; r++) {
      var row = data[r];
      var estat = iEstatus>-1 ? String(row[iEstatus]||'').trim().toLowerCase() : '';
      if (estat === 'cancelada') continue;
      var pagado = iPag>-1 ? (row[iPag]===true || String(row[iPag]).toUpperCase()==='TRUE') : false;
      if (!pagado) continue;                                  // solo egresos pagados
      var f = row[iFecha];
      if (f instanceof Date) continue;
      if (String(f||'').trim() !== '') continue;              // ya tiene algo en Fecha
      // Elegir fecha
      var nueva = null, fuente = '';
      if (iVenc>-1 && row[iVenc] instanceof Date) { nueva = row[iVenc]; fuente='vencimiento'; }
      else if (iVenc>-1 && String(row[iVenc]||'').trim() && !isNaN(new Date(row[iVenc]))) { nueva = new Date(row[iVenc]); fuente='vencimiento'; }
      if (!nueva && iMes>-1) { var pm = parseMesTag(row[iMes]); if (pm) { nueva = pm; fuente='mes'; } }
      if (!nueva) { nueva = hoy; fuente='hoy'; }
      sheet.getRange(r+1, iFecha+1).setValue(nueva);
      if (iMes>-1) sheet.getRange(r+1, iMes+1).setValue(mesesOut[nueva.getMonth()] + '-' + String(nueva.getFullYear()).slice(-2));
      reparados++;
      if (filas.length < 60) filas.push({ fila:r+1, proveedor:(iProv>-1?String(row[iProv]||''):''),
        monto:(iMonto>-1?(parseFloat(String(row[iMonto]||'').replace(/[$,]/g,''))||0):0),
        fecha:nueva.getFullYear()+'-'+String(nueva.getMonth()+1).padStart(2,'0')+'-'+String(nueva.getDate()).padStart(2,'0'), fuente:fuente });
    }
    SpreadsheetApp.flush();
    return { ok:true, reparados:reparados, filas:filas, anio:anio };
  } catch(ex) { return { ok:false, error:ex.message }; }
}

/* Guardar una REFERENCIA de texto en Link Factura / Link Pago (sin archivo) —
   para cargos sin factura ni comprobante (ej. suscripciones extranjeras que
   solo cobran). Se escribe como texto plano (borra cualquier hipervínculo
   previo). Opcionalmente marca PAGADO en el mismo movimiento. Todo desde el
   sistema, sin tocar la hoja a mano. */
function guardarReferenciaEgreso(payload) {
  try {
    var anio = payload.anio || new Date().getFullYear();
    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = EGRESOS_TABS[anio] || 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(tabName) || ss.getSheets()[0];
    if (!sheet) return {ok:false, error:'Pestaña no encontrada'};
    var rowNum = parseInt(payload.rowNum);
    if (!rowNum || rowNum < 2) return {ok:false, error:'Fila inválida'};

    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){return String(h).trim().toLowerCase();});
    function findCol(sub){ for (var c=0;c<headers.length;c++){ if (headers[c].indexOf(sub)>-1) return c; } return -1; }

    var campo = payload.campo === 'pago' ? 'pago' : 'factura';
    var texto = String(payload.texto||'').trim();
    var iDoc = findCol(campo === 'pago' ? 'link pago' : 'link factura');
    if (iDoc < 0) return {ok:false, error:'No se encontró la columna Link '+(campo==='pago'?'Pago':'Factura')};

    // Texto plano — RichText sin URL, para que el lector no lo tome como archivo
    var rich = SpreadsheetApp.newRichTextValue().setText(texto).build();
    sheet.getRange(rowNum, iDoc+1).setRichTextValue(rich);

    // Marcar la casilla del documento correspondiente (FACTURACIÓN / PAGADO)
    if (payload.marcarCasilla) {
      var iChk = findCol(campo === 'pago' ? 'pagado' : 'facturaci');
      if (iChk > -1) sheet.getRange(rowNum, iChk+1).setValue(true);
    }

    try { CacheService.getScriptCache().remove('gas_egresos_v1_'+anio); } catch(e){}
    try { logAudit(payload.usuario||'sistema','Egresos','Referencia'+(campo==='pago'?'Pago':'Factura'),'Fila '+rowNum,'','',texto); } catch(ae){}
    return {ok:true, rowNum:rowNum, campo:campo, texto:texto};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function deleteEgreso(payload) {
  try {
    var anio = payload.anio || new Date().getFullYear();
    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = EGRESOS_TABS[anio] || 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === tabName) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'Pestaña ' + tabName + ' no encontrada'};

    var rowNum = parseInt(payload.rowNum);
    if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) return {ok:false, error:'Fila inválida'};

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
    function col(name) {
      var lc = name.toLowerCase();
      for (var c=0;c<headers.length;c++) { if(headers[c].toLowerCase().indexOf(lc)>-1) return c; }
      return -1;
    }
    var rowVals = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    var iProveedor=col('proveedor'), iEgresos=col('egresos');
    var proveedor = iProveedor>-1 ? String(rowVals[iProveedor]||'') : '';
    var monto = iEgresos>-1 ? rowVals[iEgresos] : '';

    sheet.deleteRow(rowNum);

    try { logAudit(payload.usuario||'sistema', 'Egresos', 'Borrar', 'Fila '+rowNum, proveedor, '$'+monto, '—'); } catch(ae){}
    try { CacheService.getScriptCache().remove('gas_egresos_v1_' + anio); } catch(e) {}
    return {ok:true, rowNum:rowNum};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}
var PAC_COL_LISTA = 10; // columna J (1-indexed) = Lista de Precios

function _readFromBDIngresos(sheet) {
  var raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return {view:'ingresos',rows:[],resumenMensual:[],totalAnual:0,totalPagado:0,numOperaciones:0,ticketPromedio:0,topCategorias:[],topFormasPago:[],totalRows:0,headers:[]};

  function num(v){if(typeof v==='number')return v;var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}
  function dt(v){if(!v)return'';if(v instanceof Date)return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');return String(v);}

  // FacturaRFC/FacturaUUID/PagosDetalle se agregaron después (columnas dinámicas, no posición fija)
  var hdrs0 = raw[0]||[];
  var idxFacRFC = -1, idxFacUUID = -1, idxPagosDet = -1;
  for (var hci=0; hci<hdrs0.length; hci++) {
    var h0=String(hdrs0[hci]).trim();
    if (h0==='FacturaRFC') idxFacRFC=hci;
    if (h0==='FacturaUUID') idxFacUUID=hci;
    if (h0==='PagosDetalle') idxPagosDet=hci;
  }

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
      razonSocial: String(r[23]||''),
      facturaRFC: idxFacRFC>-1 ? String(r[idxFacRFC]||'') : '',
      facturaUUID: idxFacUUID>-1 ? String(r[idxFacUUID]||'') : '',
      pagosDetalle: (function(){
        if (idxPagosDet < 0) return null;
        var s = String(r[idxPagosDet]||'').trim();
        if (!s) return null;
        try { var arr = JSON.parse(s); return (arr && arr.length) ? arr : null; } catch(e){ return null; }
      })(),
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

    // Leer siempre de BD_Ingresos (fuente única)
    var bdSheet = null;
    var allSheets = ss.getSheets();
    for (var bi = 0; bi < allSheets.length; bi++) {
      if (allSheets[bi].getName() === BD_INGRESOS_TAB) { bdSheet = allSheets[bi]; break; }
    }
    if (bdSheet) {
      return _readFromBDIngresos(bdSheet);
    }
    // Fallback: leer de pestañas mensuales si BD_Ingresos no existe
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
   BD_PRODUCTOS + BD_PRECIOS + BD_AUDITORIA
   ══════════════════════════════════════════════════════════════ */

function setupBDProductos() {
  var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
  // BD_Productos
  if (!ss.getSheetByName('BD_Productos')) {
    var sh = ss.insertSheet('BD_Productos', 0);
    sh.getRange(1,1,1,8).setValues([['ProductoID','SKU','Descripcion','Categoria','Tipo','Notas','Activo','FechaCreacion']]);
    sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#f3f4f6');
    sh.setFrozenRows(1);
  }
  // BD_Precios
  if (!ss.getSheetByName('BD_Precios')) {
    var sh2 = ss.insertSheet('BD_Precios', 1);
    sh2.getRange(1,1,1,6).setValues([['ProductoID','VigenciaDesde','Precio','Moneda','ModificadoPor','FechaModificacion']]);
    sh2.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f4f6');
    sh2.setFrozenRows(1);
  }
  // BD_Auditoria
  if (!ss.getSheetByName('BD_Auditoria')) {
    var sh3 = ss.insertSheet('BD_Auditoria', 2);
    sh3.getRange(1,1,1,8).setValues([['Timestamp','Usuario','Modulo','Accion','Referencia','Campo','Anterior','Nuevo']]);
    sh3.getRange(1,1,1,8).setFontWeight('bold').setBackground('#f3f4f6');
    sh3.setFrozenRows(1);
  }
  return {ok:true, msg:'BD_Productos, BD_Precios y BD_Auditoria creadas'};
}

function logAudit(usuario, modulo, accion, referencia, campo, anterior, nuevo) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Auditoria');
    if (!sh) { setupBDProductos(); sh = ss.getSheetByName('BD_Auditoria'); }
    sh.appendRow([
      new Date(),
      usuario || 'sistema',
      modulo || '',
      accion || '',
      referencia || '',
      campo || '',
      String(anterior||''),
      String(nuevo||'')
    ]);
  } catch(e) { /* silencioso para no bloquear operaciones */ }
}

function _getNextProdID(sheet) {
  var lr = sheet.getLastRow();
  if (lr < 2) return 'PROD-00001';
  var last = String(sheet.getRange(lr,1).getValue()||'');
  var m = last.match(/PROD-(\d+)/);
  return 'PROD-' + String((m ? parseInt(m[1],10) : 0) + 1).padStart(5,'0');
}

function _getNextSkuConPrefijo(sheet, prefix) {
  var data = sheet.getDataRange().getValues();
  var re = new RegExp('^' + prefix + '-(\\d+)$');
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var m = String(data[i][1] || '').match(re);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return prefix + '-' + String(max + 1).padStart(4, '0');
}

/* ── Inventario dentro de BD_Productos ────────────────────────────
   Un solo catálogo: en vez de mantener una hoja aparte para el
   inventario de medicamentos, se agregan columnas de stock directo
   a BD_Productos. Cualquier producto (Medicamento, o cualquier otra
   categoría que se marque manualmente) puede ser "Inventariable" y
   entrar al sistema de movimientos/combos. Usa comparación EXACTA de
   encabezado (no substring) para no chocar con columnas ya existentes
   como UNIDAD_MP/UNIDAD_SAT al buscar "Unidad". */
var BDPROD_INV_HEADERS = ['Inventariable','Unidad','StockMinimo','StockMaximo','CostoUnitario','ProveedorPreferido','StockActual'];

function _bdProdColEnsure(sh, headerText) {
  var lastCol = sh.getLastColumn();
  var hdrs = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var want = headerText.trim().toLowerCase();
  var idx = hdrs.indexOf(want);
  if (idx > -1) return idx + 1;
  sh.getRange(1, lastCol+1).setValue(headerText);
  return lastCol + 1;
}

function _bdProdEnsureInventarioCols(sh) {
  var cols = {};
  BDPROD_INV_HEADERS.forEach(function(h){ cols[h] = _bdProdColEnsure(sh, h); });
  return cols;
}

// Categoría "Medicamento" siempre es inventariable de forma automática;
// otras categorías pueden marcarse manualmente vía el checkbox del formulario.
function _bdProdEsInventariable(body) {
  var categoria = String(body.categoria||'').trim().toLowerCase();
  if (categoria === 'medicamento') return true;
  return body.inventariable === true || body.inventariable === 'true';
}

function migrateProductos() {
  var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
  setupBDProductos();
  var prodSheet = ss.getSheetByName('BD_Productos');
  var precSheet = ss.getSheetByName('BD_Precios');

  // Leer todas las pestañas existentes (excepto BD_*)
  var allSheets = ss.getSheets();
  var skip = ['BD_Productos','BD_Precios','BD_Auditoria'];
  var seenProducts = {}; // key: descripcion|categoria → ProductoID
  var prodRows = [];
  var precRows = [];
  var prodCounter = 0;

  // Si ya hay datos, continuar desde el último ID
  var lr = prodSheet.getLastRow();
  if (lr > 1) {
    var existing = prodSheet.getRange(2,1,lr-1,4).getValues();
    existing.forEach(function(r){
      // Col: 0=ProductoID, 1=SKU, 2=Descripcion, 3=Categoria
      var key = (String(r[2]||'').trim() + '|' + String(r[3]||'').trim()).toLowerCase();
      seenProducts[key] = String(r[0]);
      var n = parseInt(String(r[0]).replace('PROD-',''),10);
      if (n > prodCounter) prodCounter = n;
    });
  }

  function num(v){var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}

  for (var si = 0; si < allSheets.length; si++) {
    var tab = allSheets[si];
    var tabName = tab.getName().trim();
    if (skip.indexOf(tabName) >= 0) continue;

    var raw = tab.getDataRange().getValues();
    if (raw.length < 2) continue;

    // Buscar columnas por header (español e inglés)
    var hdrs = raw[0].map(function(h){return String(h||'').trim().toLowerCase();});
    var iDesc = -1, iCat = -1, iNotas = -1, iTipo = -1, iPrecio = -1, i2025 = -1, i2026 = -1;
    for (var hi = 0; hi < hdrs.length; hi++) {
      var hv = hdrs[hi];
      if (iDesc<0 && (hv.indexOf('descripci')>=0||hv==='descripcion'||hv==='description'||hv==='product'||hv==='service'||hv==='nombre')) iDesc = hi;
      if (iCat<0 && (hv.indexOf('categori')>=0||hv==='category'||hv==='tipo de servicio')) iCat = hi;
      if (iNotas<0 && (hv.indexOf('notas')>=0||hv==='notes'||hv==='details'||hv==='adicional')) iNotas = hi;
      if (iTipo<0 && (hv==='tipo'||hv==='type'||hv==='tier')) iTipo = hi;
      if (iPrecio<0 && (hv==='precio'||hv==='price'||hv==='mxn'||hv==='usd'||hv==='costo')) iPrecio = hi;
      if (i2025<0 && hv==='2025') i2025 = hi;
      if (i2026<0 && hv==='2026') i2026 = hi;
    }
    // Si no tiene columna descripción, intentar col 0 o col 1
    if (iDesc < 0) { iDesc = hdrs.length > 2 ? 1 : 0; }
    // Si no tiene categoría, usar nombre de la pestaña
    var defaultCat = tabName;

    for (var ri = 1; ri < raw.length; ri++) {
      var r = raw[ri];
      var desc = String(r[iDesc]||'').trim();
      if (!desc) continue;
      var cat = iCat >= 0 ? String(r[iCat]||'').trim() : defaultCat;
      if (!cat) cat = defaultCat;
      var tipo = iTipo >= 0 ? String(r[iTipo]||'').trim() : '';
      var notas = iNotas >= 0 ? String(r[iNotas]||'').trim() : '';
      var precio2026 = iPrecio >= 0 ? num(r[iPrecio]) : (i2026 >= 0 ? num(r[i2026]) : 0);
      var precio2025 = i2025 >= 0 ? num(r[i2025]) : 0;

      if (!cat && !precio2026 && !precio2025) continue;

      var key = (desc + '|' + cat).toLowerCase();
      var prodId;
      if (seenProducts[key]) {
        prodId = seenProducts[key];
      } else {
        prodCounter++;
        prodId = 'PROD-' + String(prodCounter).padStart(5,'0');
        seenProducts[key] = prodId;
        // SKU auto: prefijo categoría (3 letras) + número secuencial
        var catPrefix = (cat||'GEN').substring(0,3).toUpperCase().replace(/[^A-Z]/g,'');
        var sku = catPrefix + '-' + String(prodCounter).padStart(4,'0');
        prodRows.push([prodId, sku, desc, cat, tipo, notas, true, new Date()]);
      }

      // Precios
      if (precio2025 > 0) {
        precRows.push([prodId, '2025-01-01', precio2025, 'MXN', 'migracion', new Date()]);
      }
      if (precio2026 > 0 && precio2026 !== precio2025) {
        precRows.push([prodId, '2026-01-01', precio2026, 'MXN', 'migracion', new Date()]);
      } else if (precio2026 > 0 && precio2025 === 0) {
        precRows.push([prodId, '2026-01-01', precio2026, 'MXN', 'migracion', new Date()]);
      }
    }
  }

  // Escribir productos nuevos
  if (prodRows.length) {
    prodSheet.getRange(prodSheet.getLastRow()+1, 1, prodRows.length, 8).setValues(prodRows);
  }
  // Escribir precios
  if (precRows.length) {
    precSheet.getRange(precSheet.getLastRow()+1, 1, precRows.length, 6).setValues(precRows);
  }

  logAudit('sistema', 'Productos', 'Migración', '', '', '', prodRows.length+' productos, '+precRows.length+' precios');

  return {ok:true, productosNuevos:prodRows.length, preciosNuevos:precRows.length, totalProductos:prodCounter};
}

function readProductos() {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    var precSheet = ss.getSheetByName('BD_Precios');
    if (!prodSheet || !precSheet) {
      setupBDProductos();
      prodSheet = ss.getSheetByName('BD_Productos');
      precSheet = ss.getSheetByName('BD_Precios');
    }

    // Columnas de inventario (Inventariable/Unidad/Stock…) — pueden no existir
    // todavía si ningún producto las ha usado; se leen por nombre exacto de
    // encabezado, no por posición fija, y toleran ausencia (default 0/false).
    var prodHdrs = prodSheet.getRange(1,1,1,prodSheet.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var iInv = prodHdrs.indexOf('inventariable'), iUnidad = prodHdrs.indexOf('unidad'),
        iStockMin = prodHdrs.indexOf('stockminimo'), iStockMax = prodHdrs.indexOf('stockmaximo'),
        iCosto = prodHdrs.indexOf('costounitario'), iProv = prodHdrs.indexOf('proveedorpreferido'),
        iStockAct = prodHdrs.indexOf('stockactual');

    // Leer productos
    var prodRaw = prodSheet.getDataRange().getValues();
    var productos = [];
    for (var i = 1; i < prodRaw.length; i++) {
      var r = prodRaw[i];
      var id = String(r[0]||'').trim();
      if (!id) continue;
      productos.push({
        id: id,
        sku: String(r[1]||''),
        descripcion: String(r[2]||''),
        categoria: String(r[3]||''),
        tipo: String(r[4]||''),
        notas: String(r[5]||''),
        activo: (r[6]===false||String(r[6]).toUpperCase()==='FALSE') ? false : true,
        precio: 0,
        precioVigencia: '',
        inventariable: iInv>-1 ? (r[iInv]===true || String(r[iInv]).toUpperCase()==='TRUE') : false,
        unidad: iUnidad>-1 ? String(r[iUnidad]||'') : '',
        stockMinimo: iStockMin>-1 ? (Number(r[iStockMin])||0) : 0,
        stockMaximo: iStockMax>-1 ? (Number(r[iStockMax])||0) : 0,
        costoUnitario: iCosto>-1 ? (Number(r[iCosto])||0) : 0,
        proveedorPreferido: iProv>-1 ? String(r[iProv]||'') : '',
        stockActual: iStockAct>-1 ? (Number(r[iStockAct])||0) : 0
      });
    }

    // Leer precios por lista y asignar el más reciente a cada producto
    var precRaw = precSheet.getDataRange().getValues();
    var precMap = {}; // prodId → {General:{precio,vig}, GM:{precio,vig}, ...}
    var hoy = new Date();
    function dt(v){
      if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
      return String(v||'');
    }
    for (var pi = 1; pi < precRaw.length; pi++) {
      var pr = precRaw[pi];
      var pid = String(pr[0]||'').trim();
      var vig = dt(pr[1]);
      var precio = parseFloat(String(pr[2]||'').replace(/[$,]/g,'')) || 0;
      var lista = (pr.length > 6 && pr[6] !== '' && pr[6] !== null && pr[6] !== undefined) ? String(pr[6]).trim() : 'General';
      if (!lista) lista = 'General';
      if (!pid || !precio) continue;
      if (vig > dt(hoy)) continue;
      if (!precMap[pid]) precMap[pid] = {};
      // >= para que la última fila guardada (al fondo del sheet) gane en empate de fecha
      if (!precMap[pid][lista] || vig >= precMap[pid][lista].vig) {
        precMap[pid][lista] = {precio: precio, vig: vig};
      }
    }

    var categorias = {};
    productos.forEach(function(p) {
      p.precios = {}; // {General: 78800, GM: 65000, ...}
      if (precMap[p.id]) {
        for (var lista in precMap[p.id]) {
          p.precios[lista] = precMap[p.id][lista].precio;
        }
        // Default: General, luego la primera que exista
        var gen = precMap[p.id]['General'];
        if (gen) { p.precio = gen.precio; p.precioVigencia = gen.vig; }
        else {
          var firstKey = Object.keys(precMap[p.id])[0];
          p.precio = precMap[p.id][firstKey].precio;
          p.precioVigencia = precMap[p.id][firstKey].vig;
        }
      }
      if (!categorias[p.categoria]) categorias[p.categoria] = 0;
      categorias[p.categoria]++;
    });

    var catArr = [];
    for (var c in categorias) catArr.push({nombre:c, count:categorias[c]});
    catArr.sort(function(a,b){return b.count-a.count;});

    return {
      ok: true,
      productos: productos.filter(function(p){return p.activo;}),
      todosProductos: productos,
      categorias: catArr,
      total: productos.length
    };
  } catch(ex) {
    return {ok:false, error:ex.message+' (line:'+ex.lineNumber+')', productos:[], categorias:[]};
  }
}

function updateProductoSKU(productoId, nuevoSKU, usuario) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    if (!prodSheet) return {ok:false, error:'BD_Productos no encontrada'};
    var nuevoSKU_t = String(nuevoSKU||'').trim();
    if (!nuevoSKU_t) return {ok:false, error:'SKU vacío'};

    // Verificar que no exista otro producto con el mismo SKU
    var data = prodSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === nuevoSKU_t && String(data[i][0]).trim() !== productoId) {
        return {ok:false, error:'El SKU "'+nuevoSKU_t+'" ya está asignado a '+data[i][0]};
      }
    }

    // Actualizar SKU en BD_Productos
    var skuAnterior = '';
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === productoId) {
        skuAnterior = String(data[i][1]||'');
        prodSheet.getRange(i+1, 2).setValue(nuevoSKU_t); // col B = SKU
        break;
      }
    }

    logAudit(usuario||'sistema', 'Productos', 'Cambio SKU', productoId, 'SKU', skuAnterior, nuevoSKU_t);
    return {ok:true, productoId:productoId, skuAnterior:skuAnterior, skuNuevo:nuevoSKU_t};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function updateProductoID(productoIdViejo, productoIdNuevo, usuario) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    var precSheet = ss.getSheetByName('BD_Precios');
    if (!prodSheet) return {ok:false, error:'BD_Productos no encontrada'};
    var nuevoId = String(productoIdNuevo||'').trim();
    if (!nuevoId) return {ok:false, error:'ID vacío'};

    // Verificar que no exista
    var prodData = prodSheet.getDataRange().getValues();
    for (var i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]).trim() === nuevoId) {
        return {ok:false, error:'El ID "'+nuevoId+'" ya existe'};
      }
    }

    // Actualizar en BD_Productos
    for (var i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]).trim() === productoIdViejo) {
        prodSheet.getRange(i+1, 1).setValue(nuevoId);
        break;
      }
    }

    // Propagar a BD_Precios
    if (precSheet) {
      var precData = precSheet.getDataRange().getValues();
      for (var pi = 1; pi < precData.length; pi++) {
        if (String(precData[pi][0]).trim() === productoIdViejo) {
          precSheet.getRange(pi+1, 1).setValue(nuevoId);
        }
      }
    }

    logAudit(usuario||'sistema', 'Productos', 'Cambio ID', productoIdViejo, 'ProductoID', productoIdViejo, nuevoId);
    return {ok:true, anterior:productoIdViejo, nuevo:nuevoId};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

/* ── Listas de precios ──────────────────────────────────────── */
function setupBDListas() {
  var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
  var sh = ss.getSheetByName('BD_Listas');
  if (!sh) {
    sh = ss.insertSheet('BD_Listas');
    sh.getRange(1,1,1,5).setValues([['Lista','Descripcion','Moneda','Multiplicador','Activo']]);
    sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#f3f4f6');
    sh.setFrozenRows(1);
  }
  // Si está vacía (solo header), insertar defaults
  if (sh.getLastRow() < 2) {
    sh.getRange(2,1,4,5).setValues([
      ['General','Precio estándar','MXN',1,true],
      ['GrupoMedico','Grupo Médico — precios autorizados','MXN',1,true],
      ['Surrogacy','Pacientes internacionales','USD',1,true],
      ['REPROVIDA','Precios derivados REPROVIDA','MXN',1,true]
    ]);
  }
  _syncListasDropdown();
  return {ok:true, msg:'BD_Listas configurada, dropdown sincronizado'};
}

function createProducto(body) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    if (!prodSheet) return {ok:false, error:'BD_Productos no encontrada'};
    var prodId = String(body.productoId||'').trim();
    var desc   = String(body.descripcion||'').trim();
    if (!prodId) return {ok:false, error:'ID requerido'};
    if (!desc)   return {ok:false, error:'Descripción requerida'};
    var data = prodSheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][0]).trim() === prodId)
        return {ok:false, error:'Ya existe un producto con ID: '+prodId};
    }
    var activo = body.activo !== false && body.activo !== 'false';
    prodSheet.appendRow([prodId, body.sku||'', desc, body.categoria||'', body.tipo||'', body.notas||'', activo]);
    var precio = parseFloat(String(body.precio||'').replace(/[$,]/g,''))||0;
    if (precio > 0) {
      var precSheet = ss.getSheetByName('BD_Precios');
      if (precSheet) {
        var vigencia = body.vigencia || new Date().toISOString().substring(0,10);
        precSheet.appendRow([prodId, vigencia, precio, body.moneda||'MXN', body.usuario||'sistema', new Date(), body.lista||'General']);
      }
    }
    logAudit(body.usuario||'sistema','Productos','Crear',prodId,'','-',desc);
    return {ok:true, productoId:prodId};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function leerXmlFactura(fileId) {
  try {
    if (!fileId) return {ok:false, error:'fileId requerido'};
    var file;
    try { file = DriveApp.getFileById(fileId); }
    catch(fe) { return {ok:false, error:'Archivo no encontrado: '+fe.message}; }
    // El archivo vinculado puede ser el XML (vinculado vía conciliación) o un PDF/imagen
    // subido a mano ("Subir factura PDF") — solo el XML se puede parsear como CFDI.
    var mimeType = file.getMimeType();
    var fileName = file.getName();
    var esXml = mimeType === 'text/xml' || mimeType === 'application/xml' || /\.xml$/i.test(fileName);
    if (!esXml) {
      return {ok:true, tipo:'archivo', mimeType:mimeType, fileName:fileName, fileId:fileId, viewUrl:file.getUrl()};
    }
    var content = file.getBlob().getDataAsString('UTF-8');
    if (content.charCodeAt(0) === 0xFEFF) content = content.substring(1);
    content = content.replace(/^\s+/, '');
    var doc = XmlService.parse(content);
    var root = doc.getRootElement();
    var cfdiNs = root.getNamespace();
    var attr = function(el, name) { var a = el.getAttribute(name); return a ? a.getValue() : ''; };
    var child = function(el, name) { return el.getChild(name, cfdiNs); };
    var serie = attr(root,'Serie'), folio = attr(root,'Folio'), fecha = attr(root,'Fecha');
    var subTotal = parseFloat(attr(root,'SubTotal'))||0;
    var total = parseFloat(attr(root,'Total'))||0;
    var descuento = parseFloat(attr(root,'Descuento'))||0;
    var moneda = attr(root,'Moneda')||'MXN';
    var formaPago = attr(root,'FormaPago'), metodoPago = attr(root,'MetodoPago');
    var tipoCambio = parseFloat(attr(root,'TipoCambio'))||1;
    var tipoComprobante = attr(root,'TipoDeComprobante')||'I';
    var emisorEl = child(root,'Emisor');
    var emisor = {rfc:'', nombre:'', regimen:''};
    if (emisorEl) { emisor.rfc=attr(emisorEl,'Rfc'); emisor.nombre=attr(emisorEl,'Nombre'); emisor.regimen=attr(emisorEl,'RegimenFiscal'); }
    var receptorEl = child(root,'Receptor');
    var receptor = {rfc:'', nombre:'', usoCfdi:'', domFiscal:''};
    if (receptorEl) { receptor.rfc=attr(receptorEl,'Rfc'); receptor.nombre=attr(receptorEl,'Nombre'); receptor.usoCfdi=attr(receptorEl,'UsoCFDI'); receptor.domFiscal=attr(receptorEl,'DomicilioFiscalReceptor'); }
    var conceptos = [];
    var conceptosEl = child(root,'Conceptos');
    if (conceptosEl) {
      var cEls = conceptosEl.getChildren('Concepto', cfdiNs);
      for (var i=0;i<cEls.length;i++) {
        var c=cEls[i];
        conceptos.push({claveProdServ:attr(c,'ClaveProdServ'), cantidad:attr(c,'Cantidad'), claveUnidad:attr(c,'ClaveUnidad'), descripcion:attr(c,'Descripcion'), valorUnitario:parseFloat(attr(c,'ValorUnitario'))||0, importe:parseFloat(attr(c,'Importe'))||0, descuento:parseFloat(attr(c,'Descuento'))||0});
      }
    }
    var totalImpuestosTrasladados = 0;
    var impEl = child(root,'Impuestos');
    if (impEl) totalImpuestosTrasladados = parseFloat(attr(impEl,'TotalImpuestosTrasladados'))||0;
    var uuid='', fechaTimbrado='', noCertificadoSat='', selloSat='';
    var compEl = child(root,'Complemento');
    if (compEl) {
      var chs = compEl.getChildren();
      for (var ci=0;ci<chs.length;ci++) {
        if (chs[ci].getName()==='TimbreFiscalDigital') { uuid=attr(chs[ci],'UUID'); fechaTimbrado=attr(chs[ci],'FechaTimbrado'); noCertificadoSat=attr(chs[ci],'NoCertificadoSAT'); selloSat=attr(chs[ci],'SelloSAT'); break; }
      }
    }
    var lugarExpedicion = attr(root,'LugarExpedicion');
    var noCertificado = attr(root,'NoCertificado');
    var selloCfd = attr(root,'Sello');
    return {ok:true, tipo:'xml', serie:serie, folio:folio, fecha:fecha, subTotal:subTotal, descuento:descuento, total:total, moneda:moneda, formaPago:formaPago, metodoPago:metodoPago, tipoCambio:tipoCambio, tipoComprobante:tipoComprobante, emisor:emisor, receptor:receptor, conceptos:conceptos, totalImpuestosTrasladados:totalImpuestosTrasladados, uuid:uuid, fechaTimbrado:fechaTimbrado, noCertificadoSat:noCertificadoSat, selloSat:selloSat, lugarExpedicion:lugarExpedicion, noCertificado:noCertificado, selloCfd:selloCfd, fileName:file.getName(), fileId:fileId};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function updateProducto(body) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    if (!prodSheet) return {ok:false, error:'BD_Productos no encontrada'};
    var prodId = String(body.productoId||'').trim();
    if (!prodId) return {ok:false, error:'ProductoID requerido'};
    var data = prodSheet.getDataRange().getValues();
    var found = -1;
    for (var i=1;i<data.length;i++) {
      if (String(data[i][0]).trim() === prodId) { found = i+1; break; }
    }
    if (found < 0) return {ok:false, error:'Producto no encontrado: '+prodId};
    var oldDesc = String(data[found-1][2]||'');
    // Actualizar campos (col: 0=ID, 1=SKU, 2=Desc, 3=Cat, 4=Tipo, 5=Notas, 6=Activo)
    if (body.sku !== undefined) prodSheet.getRange(found, 2).setValue(body.sku);
    if (body.descripcion !== undefined) prodSheet.getRange(found, 3).setValue(body.descripcion);
    if (body.categoria !== undefined) prodSheet.getRange(found, 4).setValue(body.categoria);
    if (body.tipo !== undefined) prodSheet.getRange(found, 5).setValue(body.tipo);
    if (body.notas !== undefined) prodSheet.getRange(found, 6).setValue(body.notas);
    if (body.activo !== undefined) prodSheet.getRange(found, 7).setValue(body.activo!==false&&body.activo!=='false');

    // Campos de inventario — se tocan solo si vienen en el body (la primera
    // vez que cualquier producto los necesita, se crean las columnas). El
    // StockActual NUNCA se edita aquí directamente: solo cambia vía
    // Compras/Ajustes/Sobrante, para que el ledger de movimientos sea la
    // única fuente de verdad del saldo.
    var invFieldMap = {inventariable:'Inventariable', unidad:'Unidad', stockMinimo:'StockMinimo',
      stockMaximo:'StockMaximo', costoUnitario:'CostoUnitario', proveedorPreferido:'ProveedorPreferido'};
    var categoriaFinal = body.categoria !== undefined ? String(body.categoria) : String(data[found-1][3]||'');
    var tocaInventario = Object.keys(invFieldMap).some(function(f){ return body[f] !== undefined; })
      || categoriaFinal.trim().toLowerCase() === 'medicamento';
    if (tocaInventario) {
      var invCols = _bdProdEnsureInventarioCols(prodSheet);
      var inventariableFinal = _bdProdEsInventariable({categoria:categoriaFinal, inventariable:body.inventariable});
      prodSheet.getRange(found, invCols.Inventariable).setValue(inventariableFinal);
      if (body.unidad !== undefined) prodSheet.getRange(found, invCols.Unidad).setValue(String(body.unidad||''));
      if (body.stockMinimo !== undefined) prodSheet.getRange(found, invCols.StockMinimo).setValue(Number(body.stockMinimo)||0);
      if (body.stockMaximo !== undefined) prodSheet.getRange(found, invCols.StockMaximo).setValue(Number(body.stockMaximo)||0);
      if (body.costoUnitario !== undefined) prodSheet.getRange(found, invCols.CostoUnitario).setValue(Number(body.costoUnitario)||0);
      if (body.proveedorPreferido !== undefined) prodSheet.getRange(found, invCols.ProveedorPreferido).setValue(String(body.proveedorPreferido||''));
    }

    // Si hay nuevo precio: borrar TODAS las filas existentes del mismo prodId+lista
    // y agregar una sola fila nueva. Esto evita la acumulación de duplicados.
    var precio = parseFloat(String(body.precio||'').replace(/[$,]/g,''))||0;
    if (precio > 0) {
      var precSheet = ss.getSheetByName('BD_Precios');
      if (precSheet) {
        var lista = body.lista || 'General';
        var vigencia = body.vigencia || new Date().toISOString().substring(0,10);
        var precData = precSheet.getDataRange().getValues();
        // Recolectar filas a borrar (en orden reverso para no desplazar índices)
        var toDelete = [];
        for (var pi = 1; pi < precData.length; pi++) {
          var plista = (precData[pi].length > 6 && precData[pi][6]) ? String(precData[pi][6]).trim() : 'General';
          if (!plista) plista = 'General';
          if (String(precData[pi][0]).trim() === prodId && plista === lista) {
            toDelete.push(pi + 1); // 1-based row number
          }
        }
        // Borrar de abajo hacia arriba para no alterar índices
        for (var di = toDelete.length - 1; di >= 0; di--) {
          precSheet.deleteRow(toDelete[di]);
        }
        // Insertar única fila con el precio actualizado
        precSheet.appendRow([prodId, vigencia, precio, body.moneda||'MXN', body.usuario||'sistema', new Date(), lista]);
      }
    }
    logAudit(body.usuario||'sistema','Productos','Editar',prodId,'Descripcion',oldDesc,body.descripcion||oldDesc);
    return {ok:true, productoId:prodId};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function readListas() {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Listas');
    if (!sh) { setupBDListas(); sh = ss.getSheetByName('BD_Listas'); }
    var raw = sh.getDataRange().getValues();
    var listas = [];
    for (var i=1;i<raw.length;i++) {
      var r=raw[i]; if(!String(r[0]||'').trim()) continue;
      listas.push({lista:String(r[0]),descripcion:String(r[1]||''),moneda:String(r[2]||'MXN'),multiplicador:Number(r[3])||1,activo:r[4]===true||String(r[4]).toUpperCase()==='TRUE'});
    }
    return {ok:true,listas:listas};
  } catch(e){return {ok:false,error:e.message,listas:[]};}
}

function saveLista(body) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Listas');
    if (!sh) { setupBDListas(); sh = ss.getSheetByName('BD_Listas'); }
    var nombre = String(body.lista||'').trim();
    if (!nombre) return {ok:false,error:'Nombre vacío'};
    // Verificar si existe para actualizar
    var data = sh.getDataRange().getValues();
    var found = -1;
    for (var i=1;i<data.length;i++) { if(String(data[i][0]).trim()===nombre){found=i+1;break;} }
    var row = [nombre, body.descripcion||'', body.moneda||'MXN', Number(body.multiplicador)||1, body.activo!==false];
    if (found>0) { sh.getRange(found,1,1,5).setValues([row]); }
    else { sh.appendRow(row); }
    // Actualizar dropdown en hoja de Pacientes
    _syncListasDropdown();
    logAudit(body.usuario||'sistema','Listas','Guardar',nombre,'','',JSON.stringify(row));
    return {ok:true,lista:nombre};
  } catch(e){return {ok:false,error:e.message};}
}

function _syncListasDropdown() {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Listas');
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    var nombres = [];
    for (var i=1;i<data.length;i++) {
      var n=String(data[i][0]||'').trim();
      if(n && (data[i][4]===true||String(data[i][4]).toUpperCase()==='TRUE')) nombres.push(n);
    }
    if (!nombres.length) return;
    // Aplicar validación a la columna J de Pacientes
    var pacSS = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var pacSh = pacSS.getSheets()[0];
    var lr = pacSh.getLastRow();
    if (lr < 2) return;
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(nombres, true).build();
    pacSh.getRange(2, PAC_COL_LISTA, lr-1, 1).setDataValidation(rule);
  } catch(e) { /* silencioso */ }
}

function migrateGMPrices() {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var gmSheet = null;
    var sheets = ss.getSheets();
    for (var i=0;i<sheets.length;i++) {
      var name = sheets[i].getName().toLowerCase();
      if (name.indexOf('autorizado')>=0 || name.indexOf('precio gm')>=0 || name.indexOf('grupomedico')>=0) { gmSheet=sheets[i]; break; }
    }
    if (!gmSheet) return {ok:false,error:'Pestaña de precios GM no encontrada'};

    var precSheet = ss.getSheetByName('BD_Precios');
    var prodSheet = ss.getSheetByName('BD_Productos');
    if (!precSheet||!prodSheet) return {ok:false,error:'BD_Precios o BD_Productos no encontrada'};

    // Leer productos existentes para buscar por descripción
    var prodData = prodSheet.getDataRange().getValues();
    var prodMap = {}; // descripcion.lower → productoId
    for (var pi=1;pi<prodData.length;pi++) {
      var desc = String(prodData[pi][2]||'').trim().toLowerCase();
      if(desc) prodMap[desc] = String(prodData[pi][0]);
    }

    var raw = gmSheet.getDataRange().getValues();
    function num(v){var n=parseFloat(String(v||'').replace(/[$,\s]/g,''));return isNaN(n)?0:n;}
    var count = 0;
    for (var ri=1;ri<raw.length;ri++) {
      var r = raw[ri];
      var desc = String(r[0]||'').trim();
      var precio = num(r[1]);
      if (!desc || !precio) continue;
      var prodId = prodMap[desc.toLowerCase()];
      if (!prodId) continue;
      precSheet.appendRow([prodId,'2026-01-01',precio,'MXN','migracion-GM',new Date(),'GM']);
      count++;
    }
    return {ok:true,migrated:count};
  } catch(e){return {ok:false,error:e.message};}
}

function saveNewProducto(body) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    var precSheet = ss.getSheetByName('BD_Precios');
    if (!prodSheet||!precSheet) { setupBDProductos(); prodSheet=ss.getSheetByName('BD_Productos'); precSheet=ss.getSheetByName('BD_Precios'); }
    var prodId = body.productoId || _getNextProdID(prodSheet);
    var desc = String(body.descripcion||'').trim();
    if (!desc) return {ok:false, error:'Descripción vacía'};
    var sku = String(body.sku||'').trim();
    prodSheet.appendRow([prodId, sku, desc, body.categoria||'', body.tipo||'', body.notas||'', body.activo!==false, new Date()]);
    var newRowNum = prodSheet.getLastRow();

    // Inventario: automático para categoría "Medicamento", opcional (checkbox)
    // para cualquier otra — así cualquier insumo/reactivo puede sumarse a
    // Combos sin necesitar un catálogo aparte.
    var inventariable = _bdProdEsInventariable(body);
    if (inventariable) {
      var invCols = _bdProdEnsureInventarioCols(prodSheet);
      prodSheet.getRange(newRowNum, invCols.Inventariable).setValue(true);
      prodSheet.getRange(newRowNum, invCols.Unidad).setValue(String(body.unidad||''));
      prodSheet.getRange(newRowNum, invCols.StockMinimo).setValue(Number(body.stockMinimo)||0);
      prodSheet.getRange(newRowNum, invCols.StockMaximo).setValue(Number(body.stockMaximo)||0);
      prodSheet.getRange(newRowNum, invCols.CostoUnitario).setValue(Number(body.costoUnitario)||0);
      prodSheet.getRange(newRowNum, invCols.ProveedorPreferido).setValue(String(body.proveedorPreferido||''));
      prodSheet.getRange(newRowNum, invCols.StockActual).setValue(0);
    }

    var precio = parseFloat(String(body.precio||'').replace(/[$,]/g,''))||0;
    if (precio>0) {
      precSheet.appendRow([prodId, body.vigencia||new Date().toISOString().substring(0,10), precio, body.moneda||'MXN', body.usuario||'sistema', new Date(), body.lista||'General']);
    }

    // Stock inicial: pasa por el motor de movimientos (no se escribe directo)
    // para que quede registrado en el ledger de dónde salió el saldo de arranque.
    var stockInicial = Number(body.stockInicial) || 0;
    if (inventariable && stockInicial > 0 && sku) {
      _registrarMovimientoInventario({
        sku: sku, nombre: desc, tipo: 'Entrada', cantidad: stockInicial, motivo: 'Alta inicial',
        referencia: '', costoUnitario: Number(body.costoUnitario)||0, usuario: body.usuario||'',
        modulo: body.categoria || 'Productos', notas: 'Alta de catálogo'
      });
    }

    logAudit(body.usuario||'sistema','Productos','Crear',prodId,'','',desc+' | $'+precio);
    return {ok:true, productoId:prodId, sku:sku};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

/* ══════════ Exportar/Importar catálogo completo (masivo) ══════════
   Permite descargar TODO el catálogo en un .xlsx editable fuera del
   sistema y volver a subirlo para actualizar muchos productos a la vez
   (precio, categoría, inventario, proveedor…) sin abrir uno por uno.
   El import empareja por ProductoID contra la hoja completa leída UNA
   sola vez en memoria y escribe con un solo setValues — así una hoja
   de ~300 productos no dispara cientos de llamadas individuales a
   Sheets (lo que arriesgaría el límite de 6 min de Apps Script). Filas
   sin ProductoID (o con uno que no existe) se dan de alta como nuevas,
   igual que el alta manual vía saveNewProducto. */
var CATALOGO_XLSX_HEADERS = ['ProductoID','SKU','Descripcion','Categoria','Tipo','Notas','Activo','Precio','Inventariable','Unidad','StockMinimo','StockMaximo','CostoUnitario','ProveedorPreferido'];

function exportarCatalogoProductos(usuario) {
  try {
    var data = readProductos();
    if (!data.ok) return {ok:false, error:data.error};
    var rows = [CATALOGO_XLSX_HEADERS];
    data.todosProductos.forEach(function(p){
      rows.push([
        p.id, p.sku, p.descripcion, p.categoria, p.tipo, p.notas,
        p.activo ? 'TRUE' : 'FALSE', p.precio||0,
        p.inventariable ? 'TRUE' : 'FALSE', p.unidad||'',
        p.stockMinimo||0, p.stockMaximo||0, p.costoUnitario||0, p.proveedorPreferido||''
      ]);
    });
    var fileName = 'Catalogo_Productos_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Mexico_City', 'yyyy-MM-dd_HHmm') + '.xlsx';
    var blob = _buildXlsxBlob(rows, 'Catalogo', fileName);
    var folder;
    try { folder = DriveApp.getFileById(PRODUCTOS_SS_ID).getParents().next(); }
    catch(e) { folder = DriveApp.getRootFolder(); }
    var file = folder.createFile(blob);
    logAudit(usuario||'sistema','Productos','ExportarCatalogo','','','',data.todosProductos.length+' productos');
    return {ok:true, url:file.getUrl(), total:data.todosProductos.length};
  } catch(ex) { return {ok:false, error:ex.message}; }
}

function importarCatalogoProductosBatch(body) {
  try {
    var productos = body.productos || [];
    if (!productos.length) return {ok:false, error:'Sin productos para importar'};
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    var precSheet = ss.getSheetByName('BD_Precios');
    if (!prodSheet || !precSheet) { setupBDProductos(); prodSheet = ss.getSheetByName('BD_Productos'); precSheet = ss.getSheetByName('BD_Precios'); }

    var invCols = _bdProdEnsureInventarioCols(prodSheet);
    var lastCol = prodSheet.getLastColumn();
    var data = prodSheet.getDataRange().getValues();

    var idxById = {};
    var maxIdNum = 0;
    for (var i = 1; i < data.length; i++) {
      var pid0 = String(data[i][0] || '').trim();
      if (pid0) {
        idxById[pid0] = i;
        var m = pid0.match(/PROD-(\d+)/);
        if (m) maxIdNum = Math.max(maxIdNum, parseInt(m[1], 10));
      }
    }

    function boolIn(v) {
      if (typeof v === 'boolean') return v;
      var s = String(v == null ? '' : v).trim().toUpperCase();
      return s === 'TRUE' || s === 'SI' || s === 'SÍ' || s === '1' || s === 'X';
    }

    var actualizados = 0, creados = 0, omitidos = 0;
    var preciosNuevos = [];
    var hoy = new Date().toISOString().substring(0, 10);
    var filasNuevas = [];

    productos.forEach(function (row) {
      var desc = String(row.descripcion || '').trim();
      if (!desc) { omitidos++; return; }
      var pid = String(row.productoId || '').trim();
      var precio = parseFloat(String(row.precio == null ? '' : row.precio).replace(/[$,]/g, '')) || 0;

      if (pid && idxById[pid] !== undefined && idxById[pid] !== null) {
        var rIdx = idxById[pid];
        if (row.sku !== undefined) data[rIdx][1] = row.sku;
        data[rIdx][2] = desc;
        if (row.categoria !== undefined) data[rIdx][3] = row.categoria;
        if (row.tipo !== undefined) data[rIdx][4] = row.tipo;
        if (row.notas !== undefined) data[rIdx][5] = row.notas;
        if (row.activo !== undefined) data[rIdx][6] = boolIn(row.activo);
        var categoriaFinal = row.categoria !== undefined ? row.categoria : data[rIdx][3];
        var inventariableFinal = _bdProdEsInventariable({ categoria: categoriaFinal, inventariable: boolIn(row.inventariable) });
        data[rIdx][invCols.Inventariable - 1] = inventariableFinal;
        if (row.unidad !== undefined) data[rIdx][invCols.Unidad - 1] = row.unidad;
        if (row.stockMinimo !== undefined) data[rIdx][invCols.StockMinimo - 1] = Number(row.stockMinimo) || 0;
        if (row.stockMaximo !== undefined) data[rIdx][invCols.StockMaximo - 1] = Number(row.stockMaximo) || 0;
        if (row.costoUnitario !== undefined) data[rIdx][invCols.CostoUnitario - 1] = Number(row.costoUnitario) || 0;
        if (row.proveedorPreferido !== undefined) data[rIdx][invCols.ProveedorPreferido - 1] = row.proveedorPreferido;
        actualizados++;
        if (precio > 0) preciosNuevos.push([pid, hoy, precio, 'MXN', body.usuario || 'sistema', new Date(), 'General']);
      } else {
        maxIdNum++;
        var newId = 'PROD-' + String(maxIdNum).padStart(5, '0');
        var newRow = new Array(lastCol).fill('');
        newRow[0] = newId;
        newRow[1] = row.sku || '';
        newRow[2] = desc;
        newRow[3] = row.categoria || '';
        newRow[4] = row.tipo || '';
        newRow[5] = row.notas || '';
        newRow[6] = row.activo !== undefined ? boolIn(row.activo) : true;
        newRow[7] = new Date();
        var inventariableNew = _bdProdEsInventariable({ categoria: row.categoria, inventariable: boolIn(row.inventariable) });
        newRow[invCols.Inventariable - 1] = inventariableNew;
        newRow[invCols.Unidad - 1] = row.unidad || '';
        newRow[invCols.StockMinimo - 1] = Number(row.stockMinimo) || 0;
        newRow[invCols.StockMaximo - 1] = Number(row.stockMaximo) || 0;
        newRow[invCols.CostoUnitario - 1] = Number(row.costoUnitario) || 0;
        newRow[invCols.ProveedorPreferido - 1] = row.proveedorPreferido || '';
        newRow[invCols.StockActual - 1] = 0;
        filasNuevas.push(newRow);
        idxById[newId] = null;
        creados++;
        if (precio > 0) preciosNuevos.push([newId, hoy, precio, 'MXN', body.usuario || 'sistema', new Date(), 'General']);
      }
    });

    if (data.length > 1) prodSheet.getRange(1, 1, data.length, lastCol).setValues(data);
    if (filasNuevas.length) prodSheet.getRange(prodSheet.getLastRow() + 1, 1, filasNuevas.length, lastCol).setValues(filasNuevas);
    if (preciosNuevos.length) precSheet.getRange(precSheet.getLastRow() + 1, 1, preciosNuevos.length, preciosNuevos[0].length).setValues(preciosNuevos);

    logAudit(body.usuario || 'sistema', 'Productos', 'ImportarCatalogoLote', '', '', '', actualizados + ' actualizados, ' + creados + ' creados, ' + omitidos + ' omitidos');
    return { ok: true, actualizados: actualizados, creados: creados, omitidos: omitidos };
  } catch (ex) { return { ok: false, error: ex.message + ' (line:' + ex.lineNumber + ')' }; }
}

function listaPacientesAll() {
  try {
    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    var hdrs = (data[0] || []).map(function(h){ return String(h).trim(); });
    var idxRS = _pacColIdx(hdrs,'Razon Social'), idxRFC = _pacColIdx(hdrs,'RFC'),
        idxCP = _pacColIdx(hdrs,'Codigo Postal'), idxUso = _pacColIdx(hdrs,'Uso CFDI'),
        idxReg = _pacColIdx(hdrs,'Regimen Fiscal'), idxFP = _pacColIdx(hdrs,'Forma de Pago Habitual');
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var nombre = String(data[i][1] || '').trim();
      if (!nombre) continue;
      var lista  = String(data[i][PAC_COL_LISTA - 1] || '').trim() || 'General';
      var email  = String(data[i][3] || '').trim();
      result.push({
        id: String(data[i][0]||'').trim(), nombre: nombre, lista: lista, email: email,
        razonSocial: idxRS>-1 ? String(data[i][idxRS]||'') : '',
        rfc: idxRFC>-1 ? String(data[i][idxRFC]||'') : '',
        codigoPostal: idxCP>-1 ? String(data[i][idxCP]||'') : '',
        usoCfdi: idxUso>-1 ? String(data[i][idxUso]||'') : '',
        regimenFiscal: idxReg>-1 ? String(data[i][idxReg]||'') : '',
        formaPagoHabitual: idxFP>-1 ? String(data[i][idxFP]||'') : ''
      });
    }
    return { ok: true, pacientes: result };
  } catch(e) { return { ok: false, error: e.message, pacientes: [] }; }
}

function readPacienteLista(pacienteNombre) {
  try {
    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      var nombre = String(data[i][1]||'').trim();
      if (nombre.toLowerCase() === String(pacienteNombre||'').trim().toLowerCase()) {
        return {ok:true, paciente:nombre, lista:String(data[i][PAC_COL_LISTA-1]||'General').trim()||'General'};
      }
    }
    return {ok:true, paciente:pacienteNombre, lista:'General'};
  } catch(e) { return {ok:false, error:e.message, lista:'General'}; }
}

/* Ficha completa de un paciente por nombre (o correo) — la fila entera
   keyed por encabezado + _rowNum, para poder EDITARLO desde el formulario
   de ingresos con el mismo flujo del catálogo (openEditModal). */
function readPacienteFull(query) {
  try {
    var q = String(query||'').trim().toLowerCase();
    if (!q) return {ok:false, error:'Falta el nombre del paciente'};
    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h).trim(); });
    // Localizar columnas de nombre y email (por si el usuario tecleó el correo)
    var iNombre = 1, iEmail = -1;
    for (var c=0;c<headers.length;c++){
      var hl = headers[c].toLowerCase();
      if (hl.indexOf('nombre')>-1 && iNombre===1) iNombre = c;
      if (hl.indexOf('mail')>-1 || hl==='e-mail' || hl==='email') iEmail = c;
    }
    for (var i=1;i<data.length;i++){
      var nombre = String(data[i][iNombre]||'').trim().toLowerCase();
      var email = iEmail>-1 ? String(data[i][iEmail]||'').trim().toLowerCase() : '';
      if (nombre === q || (email && email === q)) {
        var row = { _rowNum: i+1 };
        for (var k=0;k<headers.length;k++){ if (headers[k]) row[headers[k]] = data[i][k]; }
        // Normalizar fecha a yyyy-MM-dd si viene como Date
        for (var kk in row){
          if (row[kk] instanceof Date) row[kk] = Utilities.formatDate(row[kk], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return {ok:true, row:row};
      }
    }
    return {ok:false, error:'Ese paciente no está en el catálogo. Regístralo con el botón + antes de editarlo.'};
  } catch(e) { return {ok:false, error:e.message}; }
}

function saveProductoPrecio(productoId, nuevoPrecio, vigenciaDesde, usuario) {
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var precSheet = ss.getSheetByName('BD_Precios');
    if (!precSheet) { setupBDProductos(); precSheet = ss.getSheetByName('BD_Precios'); }

    var precio = parseFloat(String(nuevoPrecio||'').replace(/[$,]/g,'')) || 0;
    if (!precio) return {ok:false, error:'Precio inválido'};

    // Obtener precio anterior para auditoría
    var precRaw = precSheet.getDataRange().getValues();
    var precioAnterior = 0;
    for (var i = 1; i < precRaw.length; i++) {
      if (String(precRaw[i][0]) === productoId) {
        precioAnterior = parseFloat(String(precRaw[i][2]||'').replace(/[$,]/g,'')) || 0;
      }
    }

    precSheet.appendRow([
      productoId,
      vigenciaDesde || new Date().toISOString().substring(0,10),
      precio,
      'MXN',
      usuario || 'sistema',
      new Date()
    ]);

    logAudit(usuario||'sistema', 'Productos', 'Cambio precio', productoId, 'Precio',
      precioAnterior ? '$'+precioAnterior : '', '$'+precio);

    return {ok:true, productoId:productoId, nuevoPrecio:precio};
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
  'Observaciones','Factura','Poliza','USMX','CicloAltaBaja','Sucursal','ArchivoURL','RazonSocial'];

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

var INGRESOS_FOLDER_FACTURAS = '1veQEpzQPS_5FfulHP848aTVWO8CC9X-Q'; // Hestia Fertility\01 Administración y Finanzas\01 Contabilidad\Facturación
var INGRESOS_FOLDER_PAGOS    = '1D9H3nNIrkgg2wqJtKXzhuSLDH6hIUoPk';

// Columna dinámica PagosDetalle: guarda el desglose cuando una operación se
// paga con VARIAS formas (ej. tarjeta + nota de crédito). JSON compacto
// [{"fp":"Santander","monto":9220},...] en la primera línea de la OP. La
// columna FormaPago (12) conserva la forma principal (la de mayor monto)
// para no romper filtros/reportes existentes.
function _bdIngresosEnsurePagosDetalleCol(sheet) {
  var lastCol = sheet.getLastColumn();
  var hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var idx = hdrs.indexOf('pagosdetalle');
  if (idx > -1) return idx + 1;
  sheet.getRange(1, lastCol + 1).setValue('PagosDetalle');
  sheet.getRange(1, lastCol + 1).setFontWeight('bold').setBackground('#fce7f3');
  return lastCol + 1;
}

function _bdIngresosWritePagosDetalle(sheet, rowNum, pagos) {
  if (!pagos || !pagos.length) return;
  var compact = pagos.map(function(p){ return { fp: String(p.formaPago||''), monto: Math.round((parseFloat(p.monto)||0)*100)/100 }; })
    .filter(function(p){ return p.fp && p.monto > 0; });
  if (compact.length < 2) return; // un solo pago ya queda completo en FormaPago
  var col = _bdIngresosEnsurePagosDetalleCol(sheet);
  sheet.getRange(rowNum, col).setValue(JSON.stringify(compact));
}

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
    var razonSocial = payload.razonSocial || '';
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
      // Observaciones,Factura,Poliza,USMX,CicloAltaBaja,Sucursal,ArchivoURL,RazonSocial
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
        moneda, l.ciclo || ciclo, sucursal,
        '', // ArchivoURL — se llena al subir PDF
        li===0 ? razonSocial : ''
      ]);
    }

    var startRow = sheet.getLastRow()+1;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    // Pago mixto: guardar el desglose de formas de pago en la primera línea
    try { _bdIngresosWritePagosDetalle(sheet, startRow, payload.pagos); } catch (ePag) {}

    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch(e) {}

    // Descuenta inventario si alguno de los productos vendidos tiene un combo
    // configurado — nunca debe tumbar la venta si el inventario falla.
    try { _descontarInventarioPorVenta(opId, lineas, payload.usuario); } catch (eInv) {}

    return {ok:true, op:opId, lineas:rows.length, total:totalOP};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function _tokenHasPermission(token, perm) {
  try {
    var email = verifyToken(token);
    if (!email) return false;
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var userRow = getUserRow(ss, email);
    if (!userRow) return false;
    var rol = (userRow.rol || 'viewer').toLowerCase();
    if (rol === 'admin' || rol === 'director') return true;
    var rolesSh = ss.getSheetByName('Roles');
    if (!rolesSh) return false;
    var data = rolesSh.getDataRange().getValues();
    var h = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
    var rI = h.indexOf('rol'), pI = h.indexOf('permisos_operativos');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][rI] || '').trim().toLowerCase() === rol) {
        var ops = pI > -1 ? String(data[i][pI] || '').split(',').map(function(v){ return v.trim(); }).filter(Boolean) : [];
        return ops.indexOf('*') !== -1 || ops.indexOf(perm) !== -1;
      }
    }
    return false;
  } catch(ex) { return false; }
}

function updateIngreso(payload) {
  try {
    if (!_tokenHasPermission(payload.token || '', 'editar_ingresos')) {
      return {ok:false, error:'Sin autorización para editar ingresos. Solicita al administrador el permiso editar_ingresos.'};
    }
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'BD_Ingresos no encontrada'};

    var opId = payload.opId;
    if (!opId) return {ok:false, error:'opId es requerido'};

    // Encontrar y eliminar filas existentes de este OP
    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][0]).trim() === opId) rowsToDelete.push(r + 1);
    }
    // Eliminar de abajo hacia arriba para no desplazar índices
    for (var d = 0; d < rowsToDelete.length; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }

    // Re-insertar con datos actualizados (reutilizar lógica de saveIngreso)
    var lineas   = payload.lineas || [];
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
    var razonSocial = payload.razonSocial || '';
    var facturacionChk  = payload.facturacion === true || payload.facturacion === 'true';
    var conciliacionChk = payload.conciliacion === true || payload.conciliacion === 'true';
    var contabilidadChk = payload.contabilidad === true || payload.contabilidad === 'true';

    function num(v) { var n = parseFloat(String(v||'').replace(/[$,]/g,'')); return isNaN(n)?0:n; }

    var totalPagadoForm = num(payload.pagado) || 0;

    // Primera pasada: calcular totalPagar por línea y total general
    var lineCalcs = [];
    var totalOP = 0;
    for (var li = 0; li < lineas.length; li++) {
      var l = lineas[li];
      var pvp  = num(l.pvp);
      var descPct = num(l.descuento) / 100;
      var cant = num(l.cantidad) || 1;
      var tp = pvp * cant * (1 - descPct);
      totalOP += tp;
      lineCalcs.push({ l: l, pvp: pvp, cant: cant, tp: tp });
    }

    // Segunda pasada: distribuir pagado proporcionalmente entre líneas
    var rows = [];
    var remainingPagado = totalPagadoForm;
    for (var li = 0; li < lineCalcs.length; li++) {
      var lc = lineCalcs[li];
      var pagado;
      if (li === lineCalcs.length - 1) {
        pagado = Math.max(0, remainingPagado); // última línea absorbe el resto (evita errores de redondeo)
      } else {
        pagado = totalOP > 0 ? Math.round(totalPagadoForm * lc.tp / totalOP * 100) / 100 : 0;
        remainingPagado -= pagado;
      }

      rows.push([
        opId, li+1, fecha, paciente,
        lc.l.categoria||'', lc.l.producto||'',
        lc.pvp, num(lc.l.descuento), lc.cant, lc.tp,
        pagado, li===0 ? num(payload.montoFactMes) : 0,
        formaPago,
        li===0 ? facturacionChk : false,
        li===0 ? conciliacionChk : false,
        li===0 ? contabilidadChk : false,
        li===0 ? obs : '',
        li===0 ? factura : '',
        li===0 ? poliza : '',
        moneda, lc.l.ciclo || ciclo, sucursal,
        '',
        li===0 ? razonSocial : ''
      ]);
    }

    var startRowUpd = sheet.getLastRow()+1;
    sheet.getRange(startRowUpd, 1, rows.length, rows[0].length).setValues(rows);

    // Pago mixto: re-guardar el desglose (la edición borra y re-inserta filas)
    try { _bdIngresosWritePagosDetalle(sheet, startRowUpd, payload.pagos); } catch (ePag) {}

    // Auditoría
    try {
      logAudit(payload.usuario || 'sistema', 'Ingresos', 'Editar', opId, 'Operación completa',
        rowsToDelete.length + ' líneas anteriores', rows.length + ' líneas nuevas · Total: $' + totalOP.toFixed(2));
    } catch(ae) {}

    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch(e) {}
    return {ok:true, op:opId, lineas:rows.length, total:totalOP, edited:true};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function renamePacienteIngresos(oldNombre, newNombre) {
  try {
    if (!oldNombre || !newNombre || String(oldNombre).trim() === String(newNombre).trim())
      return {ok:true, updated:0, msg:'Sin cambios'};
    var oldTrim = String(oldNombre).trim();
    var newTrim = String(newNombre).trim();
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'BD_Ingresos no encontrada'};
    var data = sheet.getDataRange().getValues();
    var PAC_COL = 4; // columna D (1-indexed) = Paciente
    var updated = 0;
    for (var r = 1; r < data.length; r++) {
      var current = String(data[r][PAC_COL - 1] || '').trim();
      if (current.toLowerCase() === oldTrim.toLowerCase()) {
        sheet.getRange(r + 1, PAC_COL).setValue(newTrim);
        updated++;
      }
    }
    try { CacheService.getScriptCache().removeAll(['gas_ingresos_v1', 'gas_pacientes_v1']); } catch(e) {}
    return {ok:true, updated:updated, oldNombre:oldTrim, newNombre:newTrim};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function uploadFile(body) {
  try {
    var tipo = body.tipo || 'factura'; // factura | pago | cotizacion
    var parentId = (tipo === 'pago') ? EGRESOS_DRIVE_PAGOS
                 : (tipo === 'cotizacion') ? EGRESOS_DRIVE_COTIZACIONES
                 : EGRESOS_DRIVE_FACTURAS;
    if (!parentId) return {ok:false, error:'Carpeta no configurada para '+tipo};
    // Organizar por <Año>\<Mes> dentro de la carpeta (Pagos / Facturas / Cotizaciones)
    var hoy = new Date();
    var folder = _getOrCreateMonthFolder(parentId, hoy.getFullYear(), hoy.getMonth());
    var fileName = body.fileName || 'archivo.pdf';
    var prefix = body.prefix || ''; // ej: "EG-563" o "CXP-75"
    var fullName = prefix ? (prefix + '_' + fileName) : fileName;
    var displayName = fileName.replace(/\.(pdf|jpe?g|png|xml)$/i, '');
    // rowNum puede venir explícito o dentro del prefix tipo "CXP-75".
    var rowNum = parseInt(body.rowNum || (String(prefix).match(/(\d+)/) || [])[1] || 0);

    // 1) Resolver hoja/fila/columna y mandar a papelera el comprobante anterior ANTES de
    //    crear el nuevo (evita acumular archivos y el sufijo "(1)" por nombre repetido).
    var sh = null, hdrs = null, iCol = -1;
    if (rowNum && rowNum > 1) {
      try {
        var ss = SpreadsheetApp.openById(body.ssId || EGRESOS_SS_2026);
        sh = ss.getSheetByName(body.sheetName || (EGRESOS_TABS[2026] || 'Egresos2026'));
        if (sh) {
          hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                   .map(function(h){ return String(h).trim().toLowerCase(); });
          if (tipo === 'cotizacion') {
            iCol = _egColEnsure(sh, 'cotiz', 'Link Cotizacion') - 1; // crea la columna si no existe
          } else {
            var want = (tipo === 'pago') ? 'link pago' : 'link factura';
            for (var c = 0; c < hdrs.length; c++) { if (hdrs[c].indexOf(want) > -1) { iCol = c; break; } }
          }
          if (iCol > -1) {
            try {
              var cellRange = sh.getRange(rowNum, iCol + 1);
              var prevRich = cellRange.getRichTextValue();
              var prevUrl = (prevRich && prevRich.getLinkUrl()) ? prevRich.getLinkUrl()
                            : String(cellRange.getValue() || '');
              var mId = prevUrl.match(/[-\w]{25,}/);
              if (mId) DriveApp.getFileById(mId[0]).setTrashed(true);
            } catch(_dup) { /* el archivo previo ya no existe: ignorar */ }
          }
        }
      } catch(_sh) { sh = null; }
    }

    // 2) Crear el nuevo archivo (ya sin el viejo del mismo nombre → nombre limpio)
    var blob = Utilities.newBlob(Utilities.base64Decode(body.base64), body.mimeType || 'application/pdf', fullName);
    var file = folder.createFile(blob);
    // setSharing puede ser lento o fallar por política de la organización — no debe
    // tumbar toda la subida si el archivo ya se creó correctamente.
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_share) {}
    var url = file.getUrl();

    // 3) Escribir el hipervínculo en la fila + (si factura) marcar Facturación
    if (sh && iCol > -1) {
      sh.getRange(rowNum, iCol + 1).setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText(displayName).setLinkUrl(url).build());
    }
    if (sh && tipo === 'factura' && hdrs) {
      var iChk = -1;
      for (var k = 0; k < hdrs.length; k++) {
        if (hdrs[k].replace(/[áàä]/g,'a').indexOf('facturaci') > -1) { iChk = k; break; }
      }
      if (iChk > -1) sh.getRange(rowNum, iChk + 1).setValue(true);
    }
    logAudit(body.usuario||'sistema', 'Upload', 'Subir '+tipo, prefix||'', 'Archivo', '', fullName);
    return {ok:true, url:url, fileName:fullName};
  } catch(ex) { return {ok:false, error:ex.message}; }
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

// Actualiza SOLO campos fiscales de un OP (Factura#, Póliza) sin tocar items ni pagos.
function updateIngresoFiscal(body) {
  try {
    var opId = String(body.opId||'').trim();
    if (!opId) return {ok:false, error:'Sin opId'};

    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'BD_Ingresos no encontrada'};

    // Cols (1-based): Factura=18, Poliza=19
    var COL_FACTURA = 18, COL_POLIZA = 19;
    var data = sheet.getDataRange().getValues();
    var updated = 0;
    for (var ri = 1; ri < data.length; ri++) {
      if (String(data[ri][0]) !== opId) continue;
      if (body.hasOwnProperty('factura')) sheet.getRange(ri+1, COL_FACTURA).setValue(String(body.factura||'').trim());
      if (body.hasOwnProperty('poliza'))  sheet.getRange(ri+1, COL_POLIZA).setValue(String(body.poliza||'').trim());
      updated++;
    }
    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch(e) {}
    logAudit(body.usuario||'sistema','Ingreso','UpdateFiscal',opId,'','',
      (body.hasOwnProperty('factura')?'factura='+body.factura+' ':'')+(body.hasOwnProperty('poliza')?'poliza='+body.poliza:''));
    return {ok:true, updated:updated};
  } catch(ex) { return {ok:false, error:ex.message}; }
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

/* ══════════════════════════════════════════════════════════════
   CONCILIACIÓN AMEX — detecta cargos AMEX sin egreso registrado
   ══════════════════════════════════════════════════════════════ */
function _amexSuggestProv(desc) {
  if (!desc) return '';
  var s = desc.trim()
    .replace(/\*.*$/, '')
    .replace(/\.COM.*$/i, '')
    .replace(/\.MX.*$/i, '')
    .replace(/\s+(MX|SA|SAS|SAPI|CV|DE CV)\s*.*$/i, '')
    .replace(/[0-9]{5,}/g, '')
    .replace(/[_\-]+$/, '')
    .trim();
  if (!s || s.length < 2) s = desc.trim().split(/[\s*]/)[0];
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

function conciliaAMEX(body) {
  try {
    var periodo = body.periodo || ''; // YYYY-MM
    function num(v){
      if(typeof v==='number') return v;
      var s=String(v||'').trim().replace(/[$\s]/g,'');
      // Detectar formato europeo: punto como miles, coma como decimal (ej: "5.217,39")
      if(/\.\d{3},\d{1,2}$/.test(s)||(/,\d{1,2}$/.test(s)&&s.indexOf('.')>-1)){
        s=s.replace(/\./g,'').replace(',','.');
      } else {
        s=s.replace(/,/g,'');
      }
      var n=parseFloat(s); return isNaN(n)?0:n;
    }
    function dt(v){if(!v)return'';if(v instanceof Date)return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');return String(v).trim().substring(0,10);}
    function parseD(s){if(!s)return null;if(s instanceof Date){var d=new Date(s);d.setHours(0,0,0,0);return d;}var d=new Date(String(s).trim());if(isNaN(d))return null;d.setHours(0,0,0,0);return d;}

    // ── 1. Leer TODOS los cargos AMEX ──────────────────────────
    var bankSS = SpreadsheetApp.openById(BANKS_SS_ID);
    var allSheets = bankSS.getSheets();
    var amexSheet = null;
    for (var si = 0; si < allSheets.length; si++) {
      if (allSheets[si].getSheetId() === BANKS_GID.amex) { amexSheet = allSheets[si]; break; }
    }
    if (!amexSheet) return {ok:false, error:'No se encontró la pestaña AMEX en el spreadsheet de bancos'};

    var amexRaw = amexSheet.getDataRange().getValues();
    var amexMovs = [];
    for (var i = 1; i < amexRaw.length; i++) {
      var r = amexRaw[i];
      var monto = num(r[1]); // col B = monto; positivo = cargo
      if (monto <= 0) continue;
      var fecha = dt(r[0]);
      if (!fecha) continue;
      if (periodo && fecha.substring(0,7) !== periodo) continue;
      amexMovs.push({
        rowNum: i+1, fecha: fecha, monto: monto,
        saldo: num(r[2]), referencia: String(r[3]||'').trim(),
        usd: num(r[4]), tipoCambio: num(r[5]),
        notas: String(r[6]||''), poliza: String(r[7]||'')
      });
    }

    // ── 2. Leer Egresos con FormaPago=AMEX de VARIOS años, SIN filtrar por mes ──
    //     Un cargo del periodo puede haberse registrado en otro mes (posteo vs captura):
    //     buscamos en todo el año y en los años vecinos para NO marcarlo como faltante.
    var yBase = periodo ? parseInt(periodo.substring(0,4)) : new Date().getFullYear();
    var egresosAMEX = [];
    [yBase-1, yBase, yBase+1].forEach(function(yy){
      var egSSId = EGRESOS_IDS[yy], egTab = EGRESOS_TABS[yy];
      if (!egSSId || !egTab) return;
      var egSh; try { egSh = SpreadsheetApp.openById(egSSId).getSheetByName(egTab); } catch(e){ egSh = null; }
      if (!egSh) return;
      var egRaw = egSh.getDataRange().getValues();
      for (var i = 1; i < egRaw.length; i++) {
        var r = egRaw[i];
        if (String(r[16]||'').trim().toUpperCase() !== 'AMEX') continue; // col Q = FormaPago
        var fecha = dt(r[1]); // col B = Fecha
        if (!fecha) continue;
        egresosAMEX.push({
          anio: yy, rowNum: i+1, fecha: fecha,
          monto: Math.abs(num(r[9])), // col J = Egresos
          proveedor: String(r[4]||'').trim(),
          concepto:  String(r[8]||'').trim(),
          used: false
        });
      }
    });

    // ── 3. Cruce 1:1 — cada egreso se consume UNA sola vez (evita duplicar) ──
    //     Un cargo es "gap" solo si NO existe ningún egreso AMEX libre con el mismo
    //     importe en el rango de años. Si hay varios candidatos, gana el de fecha más
    //     cercana. Los que casan en otro mes se marcan (otroMes) para no re-registrarlos.
    var TOL = 0.5; // pesos — tolerancia por redondeo de centavos
    amexMovs.sort(function(a,b){ return a.fecha < b.fecha ? -1 : (a.fecha > b.fecha ? 1 : 0); });
    var gaps = [], conciliados = [], otroMesCount = 0;
    for (var ai = 0; ai < amexMovs.length; ai++) {
      var amov = amexMovs[ai], amovDate = parseD(amov.fecha);
      var cand = [];
      for (var ei = 0; ei < egresosAMEX.length; ei++) {
        var egr = egresosAMEX[ei];
        if (egr.used) continue;
        if (Math.abs(amov.monto - egr.monto) > TOL) continue;
        cand.push(egr);
      }
      if (cand.length) {
        cand.sort(function(a,b){ return Math.abs(amovDate-parseD(a.fecha)) - Math.abs(amovDate-parseD(b.fecha)); });
        var best = cand[0]; best.used = true;
        var dd = Math.round(Math.abs((amovDate - parseD(best.fecha)) / 86400000));
        var otroMes = best.fecha.substring(0,7) !== amov.fecha.substring(0,7);
        if (otroMes) otroMesCount++;
        conciliados.push({
          fecha: amov.fecha, monto: amov.monto, referencia: amov.referencia,
          usd: amov.usd||0, tipoCambio: amov.tipoCambio||0,
          egProveedor: best.proveedor, egConcepto: best.concepto, egFecha: best.fecha,
          difDias: dd, otroMes: otroMes
        });
      } else {
        gaps.push({
          rowNum: amov.rowNum, fecha: amov.fecha, monto: amov.monto,
          referencia: amov.referencia, usd: amov.usd, tipoCambio: amov.tipoCambio,
          proveedorSugerido: _amexSuggestProv(amov.referencia)
        });
      }
    }

    return {ok:true, gaps:gaps, conciliados:conciliados, conciliadosOtroMes:otroMesCount,
            totalAMEX:amexMovs.length, totalEgresos:egresosAMEX.length};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function readAmexCorte(body) {
  try {
    function num(v){ if(typeof v==='number') return v; var s=String(v||'').trim().replace(/[$\s,]/g,''); var n=parseFloat(s); return isNaN(n)?0:n; }
    function dt(v){ if(!v)return''; if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0'); return String(v).trim().substring(0,10); }
    function fmtD(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }

    var hoy = new Date(); hoy.setHours(0,0,0,0);
    var yr = hoy.getFullYear(), mo = hoy.getMonth(), dd = hoy.getDate();

    // Ciclo: 26 del mes anterior → 25 del mes actual
    // Si día >= 26: el ciclo que acaba de cerrarse fue 26/(mo-1) → 25/mo; el abierto es 26/mo → 25/(mo+1)
    // Si día <= 25: el cerrado fue 26/(mo-2) → 25/(mo-1);            el abierto es 26/(mo-1) → 25/mo
    var ultimoInicio, ultimoCorte, actualInicio, actualCorte;
    if (dd >= 26) {
      ultimoInicio = new Date(yr, mo-1, 26);
      ultimoCorte  = new Date(yr, mo,   25);
      actualInicio = new Date(yr, mo,   26);
      actualCorte  = new Date(yr, mo+1, 25);
    } else {
      ultimoInicio = new Date(yr, mo-2, 26);
      ultimoCorte  = new Date(yr, mo-1, 25);
      actualInicio = new Date(yr, mo-1, 26);
      actualCorte  = new Date(yr, mo,   25);
    }

    var ultimoInicioStr = fmtD(ultimoInicio), ultimoCorteStr = fmtD(ultimoCorte);
    var actualInicioStr = fmtD(actualInicio), actualCorteStr = fmtD(actualCorte);

    var bankSS = SpreadsheetApp.openById(BANKS_SS_ID);
    var amexSheet = null;
    var sheets = bankSS.getSheets();
    for (var si = 0; si < sheets.length; si++) {
      if (sheets[si].getSheetId() === BANKS_GID.amex) { amexSheet = sheets[si]; break; }
    }
    if (!amexSheet) return {ok:false, error:'No se encontró la pestaña AMEX'};

    var raw = amexSheet.getDataRange().getValues();
    var movsUltimo = [], movsActual = [];
    for (var i = 1; i < raw.length; i++) {
      var r = raw[i];
      var f = dt(r[0]);
      if (!f) continue;
      var m = num(r[1]);
      var mov = {fecha:f, monto:m, saldo:num(r[2]), referencia:String(r[3]||'').trim(), usd:num(r[4])};
      if (f >= ultimoInicioStr && f <= ultimoCorteStr) movsUltimo.push(mov);
      if (f >= actualInicioStr && f <= actualCorteStr) movsActual.push(mov);
    }

    function calcCiclo(movs) {
      var cargos = 0, pagos = 0, listaCargos = [], listaPagos = [];
      for (var j = 0; j < movs.length; j++) {
        var mv = movs[j];
        if (mv.monto > 0) { cargos += mv.monto; listaCargos.push(mv); }
        else if (mv.monto < 0) { pagos += Math.abs(mv.monto); listaPagos.push(mv); }
      }
      var saldo = Math.round((cargos - pagos) * 100) / 100;
      if (saldo < 0) saldo = 0;
      return {
        totalCargos: Math.round(cargos*100)/100,
        totalPagos:  Math.round(pagos*100)/100,
        saldoPendiente: saldo,
        cargos: listaCargos,
        pagos:  listaPagos
      };
    }

    var ultimo = calcCiclo(movsUltimo);
    var actual = calcCiclo(movsActual);

    return {
      ok: true,
      hoy: fmtD(hoy),
      ultimoCiclo: {
        inicio: ultimoInicioStr, corte: ultimoCorteStr,
        estado: ultimo.saldoPendiente < 1 ? 'pagado' : 'pendiente',
        totalCargos: ultimo.totalCargos, totalPagos: ultimo.totalPagos,
        saldoPendiente: ultimo.saldoPendiente,
        cargos: ultimo.cargos, pagos: ultimo.pagos
      },
      cicloActual: {
        inicio: actualInicioStr, corte: actualCorteStr,
        estadoCiclo: dd === 25 ? 'corte-hoy' : 'abierto',
        totalCargos: actual.totalCargos, totalPagos: actual.totalPagos,
        saldoPendiente: actual.saldoPendiente,
        cargos: actual.cargos, pagos: actual.pagos
      }
    };
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

/* ── Constantes definidas en api_config.gs (mismo proyecto GAS) ──
   SHEET_ID, AUTH_SECRET, ER_SS_ID, BANKS_SS_ID, LAB_SS_ID, MED_SS_ID,
   PAC_SS_ID, PROD_SS_ID, QX_SS_ID, ER_GID, BUDGET_GID, PL_GID,
   BANKS_GID, CAPTURA_SHEETS, SHEET_ALIASES, CAPTURA_SHEET_ID_DEFAULT
   ────────────────────────────────────────────────────────────────── */

/* ── Autenticación: helpers ──────────────────────────────────── */

/* ══════════════════════════════════════════════════════════════
   EDITAR / BORRAR INGRESOS CON SINCRONÍA BANCARIA
   ══════════════════════════════════════════════════════════════ */

function _reverseIngresoBank(opId, formaPago, fecha, monto, obs) {
  try {
    var fp  = String(formaPago || '').trim();
    var tz  = 'America/Mexico_City';
    var hoy = new Date();
    var todayStr = Utilities.formatDate(hoy, tz, 'yyyy-MM-dd');
    var mesHoy   = todayStr.substring(0, 7);
    var label    = 'REVERSO [' + opId + '] ' + String(obs || '').substring(0, 80);

    if (fp === 'Santander') {
      saveBankRow('santander', [todayStr, 0, monto, 0, label, 0, 0, '', '']);
    } else if (fp === 'TDC' || fp === 'TDD' || fp === 'AMEX' || fp === 'Transferencia') {
      saveBankRow('mercadopago', [mesHoy, todayStr, 0, 0, 0, -Math.abs(monto), 0, false, label, 'REVERSO']);
    } else if (fp === 'Efectivo') {
      var sh = getCajaChicaSheet();
      var data = sh.getDataRange().getValues();
      var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
      var iFecha  = headers.indexOf('FECHA');
      var iConc   = headers.indexOf('CONCEPTO');
      var iSalida = headers.indexOf('SALIDA');
      var targetRow = -1;
      for (var r = 1; r < data.length; r++) {
        if (!String(data[r][iFecha] || '').trim() && !String(data[r][iConc] || '').trim()) {
          targetRow = r + 1; break;
        }
      }
      if (targetRow === -1) targetRow = sh.getLastRow() + 1;
      sh.getRange(targetRow, iFecha  + 1).setValue(todayStr);
      sh.getRange(targetRow, iConc   + 1).setValue(label);
      sh.getRange(targetRow, iSalida + 1).setValue(monto);
      SpreadsheetApp.flush();
    }
    return {ok: true};
  } catch(ex) {
    Logger.log('_reverseIngresoBank error: ' + ex.message);
    return {ok: false, error: ex.message};
  }
}

// Reversa los movimientos bancarios de una OP: si tiene desglose mixto
// (PagosDetalle), reversa CADA forma por su monto — reversar todo con la
// forma principal descuadraría los bancos (ej. $17,700 de Santander cuando
// solo $9,220 entraron ahí). Nota de Crédito no toca bancos, se ignora sola.
function _reverseIngresoBankTodos(opId, hdrRow, origRows, origFP, origFecha, origPagado, obsBank) {
  var pagosArr = null;
  try {
    var hdrs = (hdrRow || []).map(function(h){ return String(h).trim().toLowerCase(); });
    var iPagosDet = hdrs.indexOf('pagosdetalle');
    if (iPagosDet > -1) {
      for (var i = 0; i < origRows.length; i++) {
        var s = String(origRows[i][iPagosDet] || '').trim();
        if (s) { pagosArr = JSON.parse(s); break; }
      }
    }
  } catch(e) { pagosArr = null; }
  if (pagosArr && pagosArr.length > 1) {
    pagosArr.forEach(function(p){
      var monto = parseFloat(p.monto) || 0;
      if (monto > 0) _reverseIngresoBank(opId, p.fp, origFecha, monto, obsBank);
    });
  } else {
    _reverseIngresoBank(opId, origFP, origFecha, origPagado, obsBank);
  }
}

/* ── Sincronía bancaria IDEMPOTENTE por OP ──────────────────────────
   En vez de reversar+re-agregar (que acumulaba filas), borra las filas
   de banco etiquetadas con [opId] y deja que se re-creen limpias; luego
   recalcula el saldo corrido (consistente con saveBankRow). Así editar N
   veces siempre deja UNA sola fila por forma de pago. Solo Santander/MP. */
function _bankSheetByKey(key) {
  var ss = SpreadsheetApp.openById(BANKS_SS_ID), sh = ss.getSheets(), gid = BANKS_GID[key];
  for (var i = 0; i < sh.length; i++) if (sh[i].getSheetId() === gid) return sh[i];
  return null;
}
function _bankObsIdx(key) { return key === 'santander' ? 4 : 8; }   // col donde va la obs con [OP-…]
/* Comisión MP (col D idx3, negativa) ya registrada para esta OP — para
   conservarla al re-crear la fila en una edición. Llamar ANTES de borrar. */
function _bankMpComisionDeOp(opId) {
  try {
    var sheet = _bankSheetByKey('mercadopago'); if (!sheet) return 0;
    var lr = sheet.getLastRow(); if (lr < 2) return 0;
    var vals = sheet.getRange(2, 1, lr - 1, Math.max(9, sheet.getLastColumn())).getValues();
    var tag = '[' + opId + ']', com = 0;
    for (var r = 0; r < vals.length; r++) { if (String(vals[r][8] || '').indexOf(tag) > -1) com += parseFloat(vals[r][3]) || 0; }
    return com;   // negativa (o 0 si no había comisión)
  } catch (e) { return 0; }
}
function _bankDeleteByOp(key, opId) {
  try {
    var sheet = _bankSheetByKey(key); if (!sheet) return 0;
    var lr = sheet.getLastRow(); if (lr < 2) return 0;
    var obsIdx = _bankObsIdx(key);
    var vals = sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).getValues();
    var tag = '[' + opId + ']', del = [];
    for (var r = 0; r < vals.length; r++) { if (String(vals[r][obsIdx] || '').indexOf(tag) > -1) del.push(r + 2); }
    for (var d = del.length - 1; d >= 0; d--) sheet.deleteRow(del[d]);
    return del.length;
  } catch (e) { Logger.log('_bankDeleteByOp ' + key + ': ' + e.message); return 0; }
}
/* Saldo de apertura ANTES de la primera fila de datos (para no perderlo al
   recalcular). Santander lleva saldo inicial; MP es cumsum puro (apertura 0).
   DEBE llamarse ANTES de borrar filas. */
function _bankOpening(key) {
  try {
    if (key !== 'santander') return 0;             // MP: saldo corrido = cumsum puro
    var sheet = _bankSheetByKey(key); if (!sheet) return 0;
    if (sheet.getLastRow() < 2) return 0;
    var r2 = sheet.getRange(2, 1, 1, 4).getValues()[0];
    var saldo2 = parseFloat(r2[3]) || 0, dep2 = parseFloat(r2[1]) || 0, ret2 = parseFloat(r2[2]) || 0;
    return saldo2 - (dep2 - ret2);                 // balance justo antes de la fila 2
  } catch (e) { return 0; }
}
function _bankRecompute(key, opening) {
  try {
    var sheet = _bankSheetByKey(key); if (!sheet) return;
    var lr = sheet.getLastRow(); if (lr < 2) return;
    if (key === 'santander') {                     // saldo col4(idx3) = apertura + Σ(dep(idx1) − ret(idx2))
      var v = sheet.getRange(2, 1, lr - 1, 4).getValues(), run = Number(opening) || 0, out = [];
      for (var i = 0; i < v.length; i++) { run += (parseFloat(v[i][1]) || 0) - (parseFloat(v[i][2]) || 0); out.push([run]); }
      sheet.getRange(2, 4, out.length, 1).setValues(out);
    } else {                                        // MP: pct col5(idx4)=1−neto/cobro ; saldo col7(idx6)=cumsum neto(idx5)
      var m = sheet.getRange(2, 1, lr - 1, 7).getValues(), runN = 0, pcts = [], sal = [];
      for (var k = 0; k < m.length; k++) { var cobro = parseFloat(m[k][2]) || 0, neto = parseFloat(m[k][5]) || 0; pcts.push([cobro !== 0 ? (1 - (neto / cobro)) : 0]); runN += neto; sal.push([runN]); }
      sheet.getRange(2, 5, pcts.length, 1).setValues(pcts);
      sheet.getRange(2, 7, sal.length, 1).setValues(sal);
    }
  } catch (e) { Logger.log('_bankRecompute ' + key + ': ' + e.message); }
}

/* Reversa SOLO los movimientos originales en Efectivo (Caja Chica); los de
   Santander/MP se manejan por _bankDeleteByOp. */
function _reverseEfectivoOnly(opId, hdrRow, origRows, origFP, origFecha, origPagado, obsBank) {
  var pagosArr = null;
  try {
    var hdrs = (hdrRow || []).map(function (h) { return String(h).trim().toLowerCase(); });
    var iP = hdrs.indexOf('pagosdetalle');
    if (iP > -1) { for (var i = 0; i < origRows.length; i++) { var s = String(origRows[i][iP] || '').trim(); if (s) { pagosArr = JSON.parse(s); break; } } }
  } catch (e) { pagosArr = null; }
  if (pagosArr && pagosArr.length) {
    pagosArr.forEach(function (p) { if (String(p.fp || '').trim() === 'Efectivo') { var m = parseFloat(p.monto) || 0; if (m > 0) _reverseIngresoBank(opId, 'Efectivo', origFecha, m, obsBank); } });
  } else if (String(origFP).trim() === 'Efectivo') {
    _reverseIngresoBank(opId, 'Efectivo', origFecha, origPagado, obsBank);
  }
}

function updateIngresoConBancos(payload) {
  try {
    if (!_tokenHasPermission(payload.token || '', 'editar_ingresos')) {
      return {ok:false, error:'Sin autorización para editar ingresos.'};
    }
    var opId = String(payload.opId || '').trim();
    if (!opId) return {ok:false, error:'opId requerido'};

    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheets = ss.getSheets();
    var sheet = null;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'BD_Ingresos no encontrada'};

    var data = sheet.getDataRange().getValues();
    var origRows = [];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === opId) origRows.push(data[r]);
    }
    if (!origRows.length) return {ok:false, error:'OP ' + opId + ' no encontrada'};

    var origFP      = String(origRows[0][12] || '');
    var origFecha   = String(origRows[0][2]  || '');
    var origPac     = String(origRows[0][3]  || '');
    var origObs     = String(origRows[0][16] || '');
    var origPagado  = origRows.reduce(function(s, r) { return s + (parseFloat(r[10]) || 0); }, 0);
    var origObsBank = (origObs ? origObs + ' · ' : '') + 'Px. ' + origPac;

    // IDEMPOTENTE: en Santander/MP se borran las filas de esta OP (se re-crean
    // limpias abajo, sin acumular reversos). Efectivo sí se reversa por su hoja.
    var _sanOpen = _bankOpening('santander');   // capturar apertura ANTES de borrar
    // Comisión MP: si el front la manda explícita (modal de edición) se usa esa;
    // si no, se conserva la ya registrada leyéndola de la fila anterior.
    var _mpCom;
    if (payload.comisionMP != null && payload.comisionMP !== '') {
      _mpCom = -Math.abs(parseFloat(payload.comisionMP) || 0);
    } else {
      _mpCom = _bankMpComisionDeOp(opId);
    }
    _bankDeleteByOp('santander', opId);
    _bankDeleteByOp('mercadopago', opId);
    _reverseEfectivoOnly(opId, data[0], origRows, origFP, origFecha, origPagado, origObsBank);

    var updateResult = updateIngreso(payload);
    if (!updateResult.ok) return updateResult;

    var payments = payload._payments || [{formaPago: payload.formaPago, monto: parseFloat(payload.pagado) || 0}];
    var newPac   = payload.paciente || '';
    var newObs   = payload.observaciones || '';
    var newFecha = payload.fecha || '';
    var mesStr   = newFecha.substring(0, 7);
    var bankObs  = (newObs ? newObs + ' · ' : '') + 'Px. ' + newPac + ' [' + opId + ']';

    var _mpComApplied = false;
    for (var pi = 0; pi < payments.length; pi++) {
      var pay = payments[pi];
      var fp  = pay.formaPago;
      var amt = parseFloat(pay.monto) || 0;
      if (!amt) continue;

      if (fp === 'Efectivo') {
        saveCajaChicaIngreso({fecha: newFecha, concepto: bankObs, entrada: amt});
      } else if (fp === 'Santander') {
        saveBankRow('santander', [newFecha, amt, 0, 0, bankObs, 0, 0, '', '']);
      } else if (fp === 'TDC' || fp === 'TDD' || fp === 'AMEX' || fp === 'Transferencia') {
        // Conserva la comisión MP ya registrada (se aplica una sola vez)
        var com = _mpComApplied ? 0 : _mpCom; _mpComApplied = true;
        saveBankRow('mercadopago', [mesStr, newFecha, amt, com, 0, amt + com, 0, false, bankObs, 'CARGO']);
      }
    }

    // Recalcula el saldo corrido tras borrar/recrear las filas de esta OP,
    // preservando el saldo de apertura de Santander.
    _bankRecompute('santander', _sanOpen);
    _bankRecompute('mercadopago', 0);

    return {ok:true, op:opId, edited:true, bankSynced:true};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

function deleteIngreso(payload) {
  try {
    if (!_tokenHasPermission(payload.token || '', 'borrar_ingresos')) {
      return {ok:false, error:'Sin autorizacion para borrar ingresos. Solicita el permiso borrar_ingresos al administrador.'};
    }
    var opId = String(payload.opId || '').trim();
    if (!opId) return {ok:false, error:'opId requerido'};

    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheets = ss.getSheets();
    var sheet = null;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; }
    }
    if (!sheet) return {ok:false, error:'BD_Ingresos no encontrada'};

    var data = sheet.getDataRange().getValues();
    var origRows = [];
    var rowNums  = [];
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][0]).trim() === opId) {
        origRows.push(data[r]);
        rowNums.push(r + 1);
      }
    }
    if (!origRows.length) return {ok:false, error:'OP ' + opId + ' no encontrada'};

    var first      = origRows[origRows.length - 1];
    var origFP     = String(first[12] || '');
    var origFecha  = String(first[2]  || '');
    var origPac    = String(first[3]  || '');
    var origObs    = String(first[16] || '');
    var origPagado = origRows.reduce(function(s, r) { return s + (parseFloat(r[10]) || 0); }, 0);
    var origObsBank = (origObs ? origObs + ' · ' : '') + 'Px. ' + origPac;

    _reverseIngresoBankTodos(opId, data[0], origRows, origFP, origFecha, origPagado, origObsBank);

    for (var d = 0; d < rowNums.length; d++) {
      sheet.deleteRow(rowNums[d]);
    }

    try {
      logAudit(payload.usuario || 'sistema', 'Ingresos', 'Borrar', opId, 'Eliminado',
        rowNums.length + ' lineas · $' + origPagado.toFixed(2), 'FP: ' + origFP);
    } catch(ae) {}

    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch(e) {}
    return {ok:true, op:opId, deleted:true, lineas:rowNums.length};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}

/* ══════════════════════════════════════════════════════════════
   CONFIG_DROPDOWNS — Opciones de formularios editables desde
   Panel de Control. Almacenadas en hoja "Config_Dropdowns"
   del spreadsheet principal (SHEET_ID).
   Columnas: Seccion | Campo | Etiqueta | Valores (|separado) | Activo
   ══════════════════════════════════════════════════════════════ */
var CFG_DD_TAB = 'Config_Dropdowns';

var CFG_DD_DEFAULTS = [
  // Seccion, Campo, Etiqueta, Valores (pipe-separated)
  ['Egresos',   'subtipo',   'Subtipo / Categoría',   'Honorarios Médicos|Honorarios Cons|Nómina|Renta|Servicios Generales|Medicamentos e Insumos|Insumos Lab|Laboratorio Externo|Marketing|Mantenimiento|Seguros|Impuestos y Contribuciones|Equipo Médico|Tecnología|Viáticos|Comisiones|Otros'],
  ['Egresos',   'contable',  'Tipo contable',          'Gasto|Costo|Crédito|Inversión'],
  ['Egresos',   'tipo',      'Tipo de egreso',         'Fijo|Variable|Extraordinario'],
  ['Egresos',   'formaPago', 'Forma de pago',          'Santander|Mercado Pago|AMEX|Efectivo|TDC|TDD|Transferencia'],
  ['CxP',       'subtipo',   'Subtipo / Categoría',   'Honorarios Médicos|Honorarios Cons|Nómina|Renta|Servicios Generales|Medicamentos e Insumos|Insumos Lab|Laboratorio Externo|Marketing|Mantenimiento|Seguros|Impuestos y Contribuciones|Equipo Médico|Tecnología|Viáticos|Comisiones|Otros'],
  ['CxP',       'formaPago', 'Forma de pago',          'Santander|Mercado Pago|AMEX|Efectivo|TDC|TDD|Transferencia'],
  ['Ingresos',  'formaPago', 'Forma de pago',          'Efectivo|Mercado Pago|AMEX|Santander|TDC|TDD|Transferencia|Cortesía'],
  ['Ingresos',  'facturacion','Facturación',            'No Factura|Factura|Factura Global|REPROVIDA|Grupo Médico'],
  ['General',   'sucursal',  'Sucursales',             'Lomas|Santa Fe|Interlomas'],
  ['General',   'moneda',    'Moneda',                 'MXN|USD|EUR'],
];

function setupConfigDropdowns() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CFG_DD_TAB);
  if (!sh) {
    sh = ss.insertSheet(CFG_DD_TAB);
    sh.getRange(1,1,1,5).setValues([['Seccion','Campo','Etiqueta','Valores','Activo']]);
    sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#c46a7a').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  // Solo insertar las filas que no existan aún
  var data = sh.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    existing[data[i][0]+'|'+data[i][1]] = true;
  }
  var toAdd = CFG_DD_DEFAULTS.filter(function(r){
    return !existing[r[0]+'|'+r[1]];
  }).map(function(r){ return [r[0], r[1], r[2], r[3], true]; });
  if (toAdd.length) sh.getRange(sh.getLastRow()+1, 1, toAdd.length, 5).setValues(toAdd);
  sh.autoResizeColumns(1,5);
  return {ok:true, msg:'Config_Dropdowns lista. Filas añadidas: '+toAdd.length};
}

function readDropdowns() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(CFG_DD_TAB);
    if (!sh) { setupConfigDropdowns(); sh = ss.getSheetByName(CFG_DD_TAB); }
    var data = sh.getDataRange().getValues();
    var result = {}; // { Seccion: { campo: { etiqueta, valores:[] } } }
    for (var i = 1; i < data.length; i++) {
      var seccion  = String(data[i][0]||'').trim();
      var campo    = String(data[i][1]||'').trim();
      var etiqueta = String(data[i][2]||'').trim();
      var valores  = String(data[i][3]||'').trim();
      var activo   = data[i][4]===true || String(data[i][4]).toUpperCase()==='TRUE';
      if (!seccion || !campo || !activo) continue;
      if (!result[seccion]) result[seccion] = {};
      result[seccion][campo] = {
        etiqueta: etiqueta,
        valores:  valores ? valores.split('|').map(function(v){return v.trim();}).filter(Boolean) : []
      };
    }
    // CxP subtipo siempre hereda de Egresos subtipo (fuente única de verdad)
    if (result['Egresos'] && result['Egresos']['subtipo']) {
      if (!result['CxP']) result['CxP'] = {};
      if (!result['CxP']['subtipo']) result['CxP']['subtipo'] = { etiqueta: 'Subtipo / Categoría', valores: [] };
      result['CxP']['subtipo'].valores = result['Egresos']['subtipo'].valores.slice();
    }
    return {ok:true, dropdowns:result};
  } catch(ex) {
    return {ok:false, error:ex.message, dropdowns:{}};
  }
}

function saveDropdownValues(body) {
  try {
    var seccion = String(body.seccion||'').trim();
    var campo   = String(body.campo  ||'').trim();
    var valores = String(body.valores||'').trim(); // pipe-separated
    if (!seccion || !campo) return {ok:false, error:'Seccion y Campo son requeridos'};

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(CFG_DD_TAB);
    if (!sh) { setupConfigDropdowns(); sh = ss.getSheetByName(CFG_DD_TAB); }

    var data  = sh.getDataRange().getValues();
    var found = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim()===seccion && String(data[i][1]).trim()===campo) {
        found = i + 1; break;
      }
    }
    if (found > 0) {
      sh.getRange(found, 4).setValue(valores); // columna D = Valores
    } else {
      var etiqueta = body.etiqueta || (campo.charAt(0).toUpperCase()+campo.slice(1));
      sh.appendRow([seccion, campo, etiqueta, valores, true]);
    }
    // Si se actualiza Egresos subtipo, sincronizar automáticamente CxP subtipo
    if (seccion === 'Egresos' && campo === 'subtipo') {
      var dataSync = sh.getDataRange().getValues();
      var cxpRow = -1;
      for (var j = 1; j < dataSync.length; j++) {
        if (String(dataSync[j][0]).trim()==='CxP' && String(dataSync[j][1]).trim()==='subtipo') {
          cxpRow = j + 1; break;
        }
      }
      if (cxpRow > 0) {
        sh.getRange(cxpRow, 4).setValue(valores);
      } else {
        sh.appendRow(['CxP', 'subtipo', 'Subtipo / Categoría', valores, true]);
      }
    }
    SpreadsheetApp.flush();
    return {ok:true, seccion:seccion, campo:campo};
  } catch(ex) {
    return {ok:false, error:ex.message};
  }
}
