/* ══════════════════════════════════════════════════════════════
   summary.gs — Summary / Estado de Resultados en vivo (Big-Four)
   ------------------------------------------------------------
   Calcula el P&L completo DIRECTO de Ingresos + Egresos para el
   rango que se elija en la app — sin depender del filtro de la hoja.
   El mapeo (qué subtipo/categoría cae en qué línea) vive en la hoja
   Summary_Config y es EDITABLE desde la página. Base devengado:
     · Egresos "Crédito" (pago AMEX del periodo anterior) NO son gasto
       nuevo — se muestran fuera del P&L como movimiento de caja.
     · "Gasto No deducible" SÍ es gasto (marcado, subtotal aparte).
     · Impuestos = línea Taxes. D&A = 0 por ahora.
   Requiere: finance.gs (readEgresosData, EGRESOS_IDS, INGRESOS_SS_*),
   core.gs/finance.gs para el wiring, SHEET_ID para la config.
   ══════════════════════════════════════════════════════════════ */

var SUMMARY_CFG_TAB = 'Summary_Config';
var SUMMARY_CFG_HEADERS = ['Fuente', 'Clave', 'Grupo', 'Linea', 'Orden', 'Flag'];
var SUMMARY_INGRESOS_IDS = null;
function _sumIngresosIds() {
  if (SUMMARY_INGRESOS_IDS) return SUMMARY_INGRESOS_IDS;
  SUMMARY_INGRESOS_IDS = {};
  SUMMARY_INGRESOS_IDS[2026] = INGRESOS_SS_ID;
  if (typeof INGRESOS_SS_2025 !== 'undefined') SUMMARY_INGRESOS_IDS[2025] = INGRESOS_SS_2025;
  if (typeof INGRESOS_SS_2024 !== 'undefined') SUMMARY_INGRESOS_IDS[2024] = INGRESOS_SS_2024;
  return SUMMARY_INGRESOS_IDS;
}

// Orden y etiqueta de los grupos del P&L
var SUMMARY_GRUPOS = {
  REVENUE:  { orden: 1, label: 'Revenue' },
  COGS:     { orden: 2, label: 'COGS' },
  OPEX:     { orden: 3, label: 'OpEx' },
  GA:       { orden: 4, label: 'G&A' },
  TAXES:    { orden: 5, label: 'Taxes' },
  EXCLUDED: { orden: 9, label: 'Fuera del P&L' }
};

function _sumNorm(s){ return String(s||'').trim().toLowerCase(); }

/* Clasificación por DEFECTO (cuando aún no está en Summary_Config).
   fuente: 'ingreso' | 'egreso'. Para egreso, clave = 'Contable|Subtipo'. */
function _summaryDefaultClass(fuente, clave) {
  if (fuente === 'ingreso') {
    var c = _sumNorm(clave);
    if (c.indexOf('surrogacy') > -1) return { grupo: 'REVENUE', linea: 'Surrogacy', orden: 2 };
    if (c.indexOf('reprovida') > -1 || c.indexOf('extern') > -1 || c.indexOf('grupo') > -1) return { grupo: 'REVENUE', linea: 'Externos', orden: 3 };
    return { grupo: 'REVENUE', linea: 'Alta', orden: 1 };
  }
  var parts = String(clave || '').split('|');
  var contable = _sumNorm(parts[0]);
  var subtipo = _sumNorm(parts[1]);
  if (contable === 'credito' || contable === 'crédito') return { grupo: 'EXCLUDED', linea: 'Pago AMEX (periodo anterior)', orden: 1 };
  if (contable === 'costo') {
    var cogsMap = { 'medicamentos':'Medicamentos','laboratorios':'Laboratorios','insumos lab':'Insumos Lab',
      'insumos qx':'Insumos Qx','comisiones':'Comisiones','honorarios':'Honorarios','honorarios cons':'Honorarios Consultas' };
    if (cogsMap[subtipo]) return { grupo:'COGS', linea:cogsMap[subtipo], orden:1 };
    return { grupo:'COGS', linea:'Other Costs', orden:9 };
  }
  // contable === 'gasto'
  if (subtipo === 'nomina' || subtipo === 'nómina') return { grupo:'GA', linea:'Nómina', orden:1 };
  if (subtipo === 'marketing') return { grupo:'OPEX', linea:'Marketing', orden:1 };
  if (subtipo === 'renta') return { grupo:'OPEX', linea:'Renta', orden:2 };
  if (subtipo === 'mtto renta') return { grupo:'OPEX', linea:'Mto Renta', orden:3 };
  if (subtipo === 'servicios' || subtipo === 'servicio') return { grupo:'OPEX', linea:'Servicios', orden:4 };
  if (subtipo === 'impuestos') return { grupo:'TAXES', linea:'Impuestos', orden:1 };
  if (subtipo === 'gasto no deducibles' || subtipo.indexOf('no deducible') > -1) return { grupo:'GA', linea:'No Deducibles', orden:8, flag:'nodeducible' };
  return { grupo:'GA', linea:'Gastos Varios', orden:5 };
}

/* Orden de los grupos de Revenue (nivel de la columna U de BD_Ingresos),
   tal como aparecen en la hoja de referencia. El grupo de cada ingreso es
   el valor de la columna U (ej. "Ciclo Alta", "Complementos", "Surrogacy"…);
   estas reglas solo definen EN QUÉ ORDEN se listan. Lo no listado va al final. */
var SUMMARY_REV_RULES = [
  { re: /alta/,                         ord: 1,  label:'Alta' },
  { re: /surrogacy|gestaci|subrog/,     ord: 2,  label:'Surrogacy' },
  { re: /externo/,                      ord: 3,  label:'Ciclos Externos' },
  { re: /baja/,                         ord: 4,  label:'Baja' },
  { re: /almacen|criopreserv/,          ord: 5,  label:'Almacenamiento' },
  { re: /estudio/,                      ord: 6,  label:'Estudios' },
  { re: /laboratorio/,                  ord: 7,  label:'Laboratorios' },
  { re: /consulta/,                     ord: 8,  label:'Consulta' },
  { re: /reprovi/,                      ord: 9,  label:'Reprovida' }
];
function _summaryRevOrden(label){
  var l = _sumNorm(label);
  for (var i=0;i<SUMMARY_REV_RULES.length;i++){ if (SUMMARY_REV_RULES[i].re.test(l)) return SUMMARY_REV_RULES[i].ord; }
  return 90; // no clasificado → al final, antes de "(Sin grupo)"
}

/* Clave normalizada (sin acentos/puntuación) para empatar contra el template. */
function _sumKey(s){
  return String(s||'').toLowerCase()
    .replace(/[áàä]/g,'a').replace(/[éèë]/g,'e').replace(/[íìï]/g,'i').replace(/[óòö]/g,'o').replace(/[úùü]/g,'u').replace(/ñ/g,'n')
    .replace(/[^a-z0-9]+/g,' ').trim();
}

/* Orden EXACTO de subgrupos (nivel 2) tal como aparecen en la hoja de referencia. */
var SUMMARY_SUBGROUP_ORDER = [
  'ciclos iniciados','alta','complementos',
  'andrology','egg donor','carrier','art lab','transfer and follow up','pgta','admin','initial consultation',
  'consulta','suplementos','estudios','laboratorios','baja','almacenamiento','ciclos externos','externos','reprovida'
];
/* Orden EXACTO de productos (nivel 3) tal como aparecen en la hoja de referencia. */
var SUMMARY_PRODUCT_ORDER = [
  // Alta · Ciclos iniciados
  'estimulacion ovarica controlada','pach estimulacion ovarica controlada',
  // Alta · Alta
  'congelacion de ovulos ef','fertilizacion in vitro ivf','pach congelamiento de ovulos','ivf transfer',
  // Alta · Complementos
  'histeroscopia','histeroscopia hospital externo','pgt a','pgt m','transferencia de embriones congelados fet',
  'consultas ultrasonidos y tomas de muestra','formacion embriones con ovulos congelados feoc','violet','magenta','biopsia testicular',
  // Surrogacy
  'sperm vitrification','andrology evaluations','ovarian stimulation and retrieval','egg donor evaluations',
  'endometrial prep','hysteroscopy','carrier evaluations','embryo formation frozen eggs','artlab cycle',
  'fet surrogacy','follow up','pgta','admin','initial consultation',
  // Ciclos Externos
  'externos','histeroscopia ext','pgt a externos','violet externos','extras externos','entidad externa',
  // Baja
  'coito programado','inseminacion intrauterina preparacion','inseminacion intrauterina iiu','monitoreo iiucp','congelamiento de esperma',
  // Almacenamiento
  'mantenimiento congelacion 12 meses','24 anualidad criopreservacion mensual','24 anualidad criopreservacion anual',
  // Estudios
  'estudios andrologia','estudios ciclo','estudios consulta',
  // Laboratorios
  'muestra de semen','compra de ovulos',
  // Consulta
  'evaluacion de fertilidad en pareja','diagnostico de fertilidad en pareja','consulta fertilidad presencial',
  'consulta fertilidad en linea','consulta de seguimiento revision resultados','consulta ginecologica presencial',
  'consulta obstetrica','consulta nutricion','consulta genetica','psicologia','sesion acupuntura rosa','consulta descontada',
  // Suplementos
  'suplementos','medicamento','vacunas gardasil'
];
function _sumOrdIn(list, label){
  var k = _sumKey(label);
  var i = list.indexOf(k);
  if (i > -1) return i;
  // empate parcial: primer elemento del template contenido en la etiqueta o viceversa
  for (var j=0;j<list.length;j++){ if (k.indexOf(list[j])>-1 || list[j].indexOf(k)>-1) return j; }
  return 999;
}

/* ── EGRESOS: subtipo → subgrupo (nivel 2) tal como agrupa la hoja de referencia.
   El nivel 1 (COGS/OpEx/G&A/Taxes) ya lo da Summary_Config; aquí solo el nivel 2. */
var SUMMARY_EG_SUBGROUP_MAP = {
  // COGS
  'honorarios':'Honorarios', 'honorarios cons':'Honorarios',
  'comisiones':'Comisiones por Venta',
  'analisis clinicos':'Estudios e Insumos de Laboratorio',
  'estudios genetica':'Estudios e Insumos de Laboratorio',
  'laboratorios':'Estudios e Insumos de Laboratorio',
  'insumos qx':'Estudios e Insumos de Laboratorio',
  'insumos lab':'Estudios e Insumos de Laboratorio',
  'gases':'Estudios e Insumos de Laboratorio',
  'renta equipo':'Estudios e Insumos de Laboratorio',
  'medicamentos':'Medicamentos',
  'software':'Servicios', 'reportes':'Servicios',
  // OpEx
  'renta':'Inmueble', 'mtto renta':'Inmueble', 'mto renta':'Inmueble',
  'valet parking':'Inmueble', 'mantenimiento':'Inmueble',
  'marketing':'Marketing y Publicidad', 'whatsapp':'Marketing y Publicidad',
  'mtto web':'Marketing y Publicidad', 'podcast':'Marketing y Publicidad', 'adds':'Marketing y Publicidad',
  'servicios':'Servicios', 'servicio':'Servicios', 'contabilidad':'Servicios', 'rpbi':'Servicios'
  // (G&A y Taxes: se agregan al recibir las 2 capturas faltantes)
};
/* Orden de subgrupos dentro de su sección. "Servicios" siempre al final de su sección. */
var SUMMARY_EG_SUBGROUP_ORDER = {
  'honorarios':1, 'comisiones por venta':2, 'estudios e insumos de laboratorio':3,
  'medicamentos':4, 'inmueble':5, 'marketing y publicidad':6, 'nomina':7, 'nómina':7,
  'gastos varios':8, 'no deducibles':9, 'impuestos':10, 'servicios':90
};
function _summaryEgSubgroup(subtipo){
  var k = _sumNorm(subtipo);
  return SUMMARY_EG_SUBGROUP_MAP[k] || (String(subtipo||'').trim() || 'Otros');
}
function _summaryEgSubgroupOrden(subgroup){
  var o = SUMMARY_EG_SUBGROUP_ORDER[_sumNorm(subgroup)];
  return (o===undefined) ? 50 : o;
}

/* ── Siembra/actualiza Summary_Config escaneando los datos reales.
   Corre UNA VEZ (o cuando quieras re-sembrar). No pisa lo que ya
   editaste: solo AGREGA las claves nuevas que falten. ─────────────── */
function setupSummaryConfig() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SUMMARY_CFG_TAB);
  if (!sh) {
    sh = ss.insertSheet(SUMMARY_CFG_TAB);
    sh.getRange(1,1,1,SUMMARY_CFG_HEADERS.length).setValues([SUMMARY_CFG_HEADERS]);
    sh.getRange(1,1,1,SUMMARY_CFG_HEADERS.length).setFontWeight('bold').setBackground('#1a252f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  // Claves ya presentes
  var data = sh.getDataRange().getValues();
  var existing = {};
  for (var i=1;i<data.length;i++){ existing[_sumNorm(data[i][0])+'::'+_sumNorm(data[i][1])] = true; }

  // Escanear claves reales de Egresos (contable|subtipo) e Ingresos (categoria)
  var claves = {}; // 'fuente::clave' -> {fuente, clave}
  var anio = new Date().getFullYear();
  try {
    var eg = readEgresosData(anio);
    (eg.rows||[]).forEach(function(r){
      var clave = String(r.contable||'')+'|'+String(r.subtipo||'');
      claves['egreso::'+clave] = { fuente:'egreso', clave:clave };
    });
  } catch(e){}
  try {
    var yid = _sumIngresosIds()[anio] || INGRESOS_SS_ID;
    var ish = SpreadsheetApp.openById(yid).getSheetByName('BD_Ingresos') || SpreadsheetApp.openById(yid).getSheets()[0];
    var idata = ish.getDataRange().getValues();
    for (var ii=1; ii<idata.length; ii++){
      var cat = String(idata[ii][4]||'').trim();
      if (cat) claves['ingreso::'+cat] = { fuente:'ingreso', clave:cat };
    }
  } catch(e2){}

  var nuevas = [];
  Object.keys(claves).forEach(function(k){
    var it = claves[k];
    if (existing[_sumNorm(it.fuente)+'::'+_sumNorm(it.clave)]) return;
    var cls = _summaryDefaultClass(it.fuente, it.clave);
    nuevas.push([it.fuente, it.clave, cls.grupo, cls.linea, cls.orden||5, cls.flag||'']);
  });
  if (nuevas.length) sh.getRange(sh.getLastRow()+1, 1, nuevas.length, SUMMARY_CFG_HEADERS.length).setValues(nuevas);
  return { ok:true, agregadas:nuevas.length, totalClaves:Object.keys(claves).length };
}

function readSummaryConfig() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(SUMMARY_CFG_TAB);
    if (!sh) { setupSummaryConfig(); sh = ss.getSheetByName(SUMMARY_CFG_TAB); }
    var data = sh.getDataRange().getValues();
    var rows = [];
    for (var i=1;i<data.length;i++){
      if (!String(data[i][1]||'').trim()) continue;
      rows.push({ fuente:String(data[i][0]||''), clave:String(data[i][1]||''), grupo:String(data[i][2]||''),
        linea:String(data[i][3]||''), orden:Number(data[i][4])||5, flag:String(data[i][5]||'') });
    }
    return { ok:true, rows:rows, grupos:SUMMARY_GRUPOS };
  } catch(ex){ return { ok:false, error:ex.message, rows:[] }; }
}

function saveSummaryConfig(body) {
  try {
    var rows = body.rows || [];
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(SUMMARY_CFG_TAB);
    if (!sh) { setupSummaryConfig(); sh = ss.getSheetByName(SUMMARY_CFG_TAB); }
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2,1,last-1,SUMMARY_CFG_HEADERS.length).clearContent();
    if (rows.length) {
      var out = rows.map(function(r){ return [r.fuente||'', r.clave||'', r.grupo||'GA', r.linea||'Gastos Varios', Number(r.orden)||5, r.flag||'']; });
      sh.getRange(2,1,out.length,SUMMARY_CFG_HEADERS.length).setValues(out);
    }
    try { logAudit(body.usuario||'sistema','Summary','GuardarConfig','','','',rows.length+' líneas'); } catch(e){}
    return { ok:true, guardadas:rows.length };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Lectura completa de Ingresos de un año (BD_Ingresos, sin recorte) ── */
function _summaryReadIngresos(anio) {
  var out = [];
  var yid = _sumIngresosIds()[anio];
  if (!yid) return out;
  var ss = SpreadsheetApp.openById(yid);
  var sh = ss.getSheetByName('BD_Ingresos') || ss.getSheets()[0];
  if (!sh) return out;
  var data = sh.getDataRange().getValues();
  function num(v){ if(typeof v==='number') return v; var n=parseFloat(String(v||'').replace(/[$,\s]/g,'')); return isNaN(n)?0:n; }
  for (var i=1;i<data.length;i++){
    var r=data[i];
    if (!String(r[0]||'').trim()) continue;
    out.push({ op:String(r[0]||''), fecha:_sumParseDate(r[2]), paciente:String(r[3]||''),
      categoria:String(r[4]||''), producto:String(r[5]||''), cantidad:num(r[8]),
      total:num(r[9]), formaPago:String(r[12]||''),
      grupoU:String(r[20]||'').trim() });   // columna U (índice 20) = grupo/categoría del reporte
  }
  return out;
}

/* Parseo robusto de fecha → 'YYYY-MM-DD' (Date, ISO, dd/mm/yyyy, dd-mm-yyyy).
   Devuelve '' si no se puede interpretar (para no descuadrar por fechas basura). */
function _sumParseDate(v){
  if (v instanceof Date && !isNaN(v)) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
  var s = String(v||'').trim(); if (!s) return '';
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);           // yyyy-mm-dd
  if (m) return m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0');
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);               // dd/mm/yyyy
  if (m) return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
  return '';
}

function _sumInRange(f, ini, fin){ return f && f>=ini && f<=fin; }
function _sumFmtDate(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }

/* ── El reporte ─────────────────────────────────────────────────── */
function readSummary(fechaInicio, fechaFin) {
  try {
    var fi = String(fechaInicio||'').substring(0,10);
    var ff = String(fechaFin||'').substring(0,10);
    if (!fi || !ff) { var hoy=new Date(); ff=_sumFmtDate(hoy); fi=ff.substring(0,8)+'01'; }

    // Periodo anterior: mismo tamaño, inmediatamente antes
    var dIni=new Date(fi+'T12:00:00'), dFin=new Date(ff+'T12:00:00');
    var dias = Math.round((dFin-dIni)/86400000)+1;
    var pFin=new Date(dIni.getTime()-86400000);
    var pIni=new Date(pFin.getTime()-(dias-1)*86400000);
    var pi=_sumFmtDate(pIni), pf=_sumFmtDate(pFin);

    var cfg = readSummaryConfig();
    var mapEg={}, mapIn={};
    (cfg.rows||[]).forEach(function(r){
      if (_sumNorm(r.fuente)==='ingreso') mapIn[_sumNorm(r.clave)] = r;
      else mapEg[_sumNorm(r.clave)] = r;
    });
    function clasifEg(clave){ return mapEg[_sumNorm(clave)] || _summaryDefaultClass('egreso', clave); }
    function clasifIn(cat){ return mapIn[_sumNorm(cat)] || _summaryDefaultClass('ingreso', cat); }

    // Acumuladores por línea (con sub-items agrupados por producto/concepto)
    var agg = {}; // 'GRUPO|Linea' -> {grupo,linea,orden,flag,actual,prev,subs:{}}
    function key(g,l){ return g+'|'+l; }
    function get(g,l,orden,flag){ var k=key(g,l); if(!agg[k]) agg[k]={grupo:g,linea:l,orden:orden||5,flag:flag||'',actual:0,prev:0,subs:{}}; return agg[k]; }
    function addSub(line, label, monto, cant, isA, drill, grupo, meta){
      label = String(label||'(sin nombre)').trim() || '(sin nombre)';
      grupo = String(grupo||'').trim();
      var sk = grupo+'||'+label;
      var s = line.subs[sk];
      if(!s){ s = line.subs[sk] = {label:label, grupo:grupo, meta:meta||'', actual:0, prev:0, cantA:0, cantP:0, rows:[], _ord:Object.keys(line.subs).length}; }
      else if(meta && !s.meta){ s.meta = meta; }
      if(isA){ s.actual += monto; s.cantA += cant; if(drill) s.rows.push(drill); }
      else { s.prev += monto; s.cantP += cant; }
    }

    // Acumulador de REVENUE en 3 niveles: Grupo (col U) → Subgrupo (categoría) → Producto → movimientos
    var revAgg = {}; // l1 -> {label,orden,actual,prev,subs:{ l2 -> {label,actual,prev,cantA,cantP,_ord,prods:{ prod -> {label,actual,prev,cantA,cantP,rows:[]} }} }}
    function getRev(l1, orden){ if(!revAgg[l1]) revAgg[l1]={label:l1,orden:orden||90,actual:0,prev:0,subs:{}}; return revAgg[l1]; }
    function addRev(l1line, l2, prod, monto, cant, isA, drill){
      l2 = String(l2||'').trim() || l1line.label;
      prod = String(prod||'(sin producto)').trim() || '(sin producto)';
      var s = l1line.subs[l2];
      if(!s){ s = l1line.subs[l2] = {label:l2, actual:0, prev:0, cantA:0, cantP:0, _ord:Object.keys(l1line.subs).length, prods:{}}; }
      var p = s.prods[prod];
      if(!p){ p = s.prods[prod] = {label:prod, actual:0, prev:0, cantA:0, cantP:0, _ord:Object.keys(s.prods).length, rows:[]}; }
      if(isA){ s.actual+=monto; s.cantA+=cant; p.actual+=monto; p.cantA+=cant; if(drill) p.rows.push(drill); }
      else { s.prev+=monto; s.cantP+=cant; p.prev+=monto; p.cantP+=cant; }
    }

    var amex = { actual:0, prev:0, rows:[] };
    var recon = { ingresosTotal:0, egresosTotal:0, egresosBruto:0, egresosCancelado:0, sinClasificar:0,
                  ingresosSinFecha:0, ingresosSinFechaMonto:0, egresosSinFecha:0 };
    var labInsMap = {}; // subtipo -> {subtipo,total,count}
    var LAB_SUBS = {'insumos lab':1,'insumos qx':1,'laboratorios':1,'medicamentos':1,'gases':1,'reportes':1,'renta equipo':1};

    // Años a leer (incluye el año del periodo anterior)
    var yFrom = parseInt(pi.substring(0,4),10), yTo = parseInt(ff.substring(0,4),10);

    // ── Egresos ──
    for (var y=yFrom; y<=yTo; y++){
      var eg; try { eg = readEgresosData(y); } catch(e){ continue; }
      (eg.rows||[]).forEach(function(r){
        var f = (r.fecha||r.vencimiento||'').substring(0,10);
        var enActual = _sumInRange(f, fi, ff), enPrev = _sumInRange(f, pi, pf);
        if (!enActual && !enPrev) return;
        var monto = Number(r.monto)||0;
        if (r.estatus === 'Cancelada'){ if(enActual) recon.egresosCancelado += monto; return; }
        // Bruto de caja del periodo: TODO lo capturado no cancelado (incluye el pago AMEX/Crédito)
        if (enActual) recon.egresosBruto += monto;
        var clave = String(r.contable||'')+'|'+String(r.subtipo||'');
        var c = clasifEg(clave);
        if (c.grupo === 'EXCLUDED') {
          if (enActual){ amex.actual+=monto; amex.rows.push({fecha:f,nombre:r.proveedor,concepto:r.concepto,monto:monto}); }
          if (enPrev) amex.prev+=monto;
          return;
        }
        // Nivel 2 (dato) = SUBGRUPO curado (Honorarios, Inmueble, Medicamentos…) desde el subtipo
        var sub2 = _summaryEgSubgroup(r.subtipo);
        var line = get(c.grupo, sub2, _summaryEgSubgroupOrden(sub2), c.flag);
        // Nivel 3 (sub-item) = PROVEEDOR, con concepto·tipo como meta
        var provLbl = String(r.proveedor||'').trim() || String(r.concepto||'').trim() || '(sin proveedor)';
        var meta = [String(r.concepto||'').trim(), String(r.subtipo||'').trim(), String(r.tipo||'').trim()].filter(Boolean).join(' · ');
        if (enActual){ line.actual+=monto; recon.egresosTotal+=monto;
          addSub(line, provLbl, monto, 1, true, {fecha:f, nombre:r.proveedor, concepto:(String(r.concepto||'')+' · '+String(r.subtipo||'')), monto:monto, subtipo:r.subtipo}, '', meta);
          var sk=_sumNorm(r.subtipo); if (LAB_SUBS[sk]){ if(!labInsMap[sk]) labInsMap[sk]={subtipo:r.subtipo,total:0,count:0}; labInsMap[sk].total+=monto; labInsMap[sk].count++; }
        }
        if (enPrev){ line.prev+=monto; addSub(line, provLbl, monto, 1, false, null, '', meta); }
      });
    }

    // ── Ingresos ──
    for (var yi=yFrom; yi<=yTo; yi++){
      var ins = _summaryReadIngresos(yi);
      ins.forEach(function(r){
        var f=(r.fecha||'').substring(0,10);
        if (!f && Number(r.total)){ recon.ingresosSinFecha++; recon.ingresosSinFechaMonto += Number(r.total)||0; }
        var enActual=_sumInRange(f, fi, ff), enPrev=_sumInRange(f, pi, pf);
        if (!enActual && !enPrev) return;
        // 3 niveles: Grupo (col U) → Subgrupo (categoría) → Producto
        var l1 = String(r.grupoU||'').trim() || String(r.categoria||'').trim() || '(Sin grupo)';
        var l2 = String(r.categoria||'').trim();
        var prod = String(r.producto||'').trim() || l2 || '(sin producto)';
        var cant = Number(r.cantidad)||0; if(!cant) cant=1;
        var line = getRev(l1, _summaryRevOrden(l1));
        if (enActual){ line.actual+=r.total; recon.ingresosTotal+=r.total;
          addRev(line, l2, prod, r.total, cant, true, {fecha:f, nombre:r.paciente, concepto:r.producto+' · '+r.categoria, monto:r.total, cantidad:cant});
        }
        if (enPrev){ line.prev+=r.total; addRev(line, l2, prod, r.total, cant, false, null); }
      });
    }

    // Totales por grupo
    function grpTot(g,campo){ var t=0; Object.keys(agg).forEach(function(k){ if(agg[k].grupo===g) t+=agg[k][campo]||0; }); return t; }
    var revA=0, revP=0; Object.keys(revAgg).forEach(function(k){ revA+=revAgg[k].actual; revP+=revAgg[k].prev; });
    var cogsA=grpTot('COGS','actual'), cogsP=grpTot('COGS','prev');
    var opexA=grpTot('OPEX','actual'), opexP=grpTot('OPEX','prev');
    var gaA=grpTot('GA','actual'), gaP=grpTot('GA','prev');
    var taxA=grpTot('TAXES','actual'), taxP=grpTot('TAXES','prev');
    var da=0;
    var gpA=revA-cogsA, gpP=revP-cogsP;
    var ccA=gpA-opexA, ccP=gpP-opexP;
    var ebitdaA=ccA-gaA, ebitdaP=ccP-gaP;
    var ebitA=ebitdaA-da, ebitP=ebitdaP-da;
    var netA=ebitA-taxA, netP=ebitP-taxP;

    // Ensamblar líneas ordenadas
    function lineasDe(g){
      var arr=[]; Object.keys(agg).forEach(function(k){ if(agg[k].grupo===g) arr.push(agg[k]); });
      arr.sort(function(a,b){ return (a.orden-b.orden) || (a.linea<b.linea?-1:1); });
      return arr.map(function(l){
        // Sub-items (proveedores) ordenados por monto desc; ocultar los de $0
        var subsArr = Object.keys(l.subs).map(function(k){ return l.subs[k]; })
          .filter(function(s){ return Math.abs(s.actual)>0.005 || Math.abs(s.prev)>0.005; })
          .sort(function(a,b){ return b.actual-a.actual; });
        var subs = subsArr.map(function(s){ return { label:s.label, meta:s.meta||'', cantidad:s.cantA, cantidadPrev:s.cantP,
            actual:s.actual, prev:s.prev,
            rows:s.rows.sort(function(a,b){ return b.monto-a.monto; }) }; });
        return { tipo:'dato', grupo:g, linea:l.linea, label:l.linea, flag:l.flag,
          actual:l.actual, prev:l.prev, subitems:subs };
      });
    }
    // REVENUE en 3 niveles: Grupo → Subgrupo → Producto (con aplanado si hay 1 solo subgrupo)
    function revLineas(){
      var arr = Object.keys(revAgg).map(function(k){ return revAgg[k]; });
      arr.sort(function(a,b){ return (a.orden-b.orden) || (b.actual-a.actual); });
      return arr.map(function(l1){
        var subsArr = Object.keys(l1.subs).map(function(k){ return l1.subs[k]; })
          .sort(function(a,b){ var oa=_sumOrdIn(SUMMARY_SUBGROUP_ORDER,a.label), ob=_sumOrdIn(SUMMARY_SUBGROUP_ORDER,b.label);
            return (oa-ob) || (b.actual-a.actual); });
        function prodsOf(s){
          return Object.keys(s.prods).map(function(k){ return s.prods[k]; })
            .sort(function(a,b){ var oa=_sumOrdIn(SUMMARY_PRODUCT_ORDER,a.label), ob=_sumOrdIn(SUMMARY_PRODUCT_ORDER,b.label);
              return (oa-ob) || (b.actual-a.actual); })
            .map(function(p){ return { label:p.label, cantidad:p.cantA, actual:p.actual, prev:p.prev,
              rows:p.rows.sort(function(a,b){ return b.monto-a.monto; }) }; });
        }
        var subitems;
        if (subsArr.length <= 1){
          // Un solo subgrupo → mostrar productos directos (nivel 2 = producto)
          subitems = subsArr.length ? prodsOf(subsArr[0]) : [];
        } else {
          // Varios subgrupos → nivel 2 = subgrupo, nivel 3 = producto
          subitems = subsArr.map(function(s){ return { label:s.label, cantidad:s.cantA,
            actual:s.actual, prev:s.prev, productos:prodsOf(s) }; });
        }
        return { tipo:'dato', grupo:'REVENUE', linea:l1.label, label:l1.label,
          actual:l1.actual, prev:l1.prev, subitems:subitems };
      });
    }
    function pct(v){ return revA>0 ? v/revA : null; }
    function sec(g,label,actual,prev){ return { tipo:'seccion', grupo:g, label:label, actual:actual, prev:prev, pct:pct(actual) }; }
    function met(label,actual,prev){ return { tipo:'metric', label:label, actual:actual, prev:prev, pct:pct(actual) }; }
    function stamp(list){ return list.map(function(x){ x.pct=pct(x.actual); return x; }); }

    var lineas = [];
    lineas.push(sec('REVENUE','Revenue',revA,revP));  lineas = lineas.concat(stamp(revLineas()));
    lineas.push(sec('COGS','COGS',cogsA,cogsP));       lineas = lineas.concat(stamp(lineasDe('COGS')));
    lineas.push(met('Gross Profit',gpA,gpP));
    lineas.push(sec('OPEX','OpEx',opexA,opexP));       lineas = lineas.concat(stamp(lineasDe('OPEX')));
    lineas.push(met('Clinic Contribution',ccA,ccP));
    lineas.push(sec('GA','G&A',gaA,gaP));              lineas = lineas.concat(stamp(lineasDe('GA')));
    lineas.push(met('EBITDA',ebitdaA,ebitdaP));
    lineas.push({tipo:'dato',grupo:'DA',label:'Depreciación & Amortización',actual:da,prev:0,rows:[],pct:pct(da)});
    lineas.push(met('EBIT',ebitA,ebitP));
    lineas.push(met('EBT',ebitA,ebitP));
    lineas.push(sec('TAXES','Taxes',taxA,taxP));       lineas = lineas.concat(stamp(lineasDe('TAXES')));
    lineas.push(met('Net Profit',netA,netP));

    var labInsumos=[]; Object.keys(labInsMap).forEach(function(k){ labInsumos.push(labInsMap[k]); });
    labInsumos.sort(function(a,b){ return b.total-a.total; });

    var egSum = cogsA+opexA+gaA+taxA;
    return {
      ok:true,
      periodo:{ inicio:fi, fin:ff }, prev:{ inicio:pi, fin:pf },
      lineas: lineas,
      metricas:{ revenue:revA, cogs:cogsA, grossProfit:gpA, opex:opexA, clinicContribution:ccA,
                 ga:gaA, ebitda:ebitdaA, da:da, ebit:ebitA, taxes:taxA, netProfit:netA },
      amexCredito: amex,
      labInsumos: labInsumos,
      reconc:{
        // Ingresos: el reporte = todo lo capturado (no hay exclusiones)
        ingresosBruto: recon.ingresosTotal,
        ingresosSum: revA,
        cuadraIngresos: Math.abs(recon.ingresosTotal-revA)<0.5,
        // Egresos: puente caja → P&L devengado
        egresosBruto: recon.egresosBruto,        // TODO lo capturado (incluye pago AMEX), sin cancelados
        creditoAmex: amex.actual,                // pago AMEX del periodo anterior (fuera del P&L)
        egresosCancelado: recon.egresosCancelado,// cancelados (no cuentan)
        egresosPL: egSum,                        // COGS+OpEx+G&A+Taxes (devengado)
        // Identidad: bruto = P&L + crédito AMEX
        cuadraEgresos: Math.abs(recon.egresosBruto - amex.actual - egSum)<0.5,
        // Alertas: filas con fecha inválida que NO entran a ningún periodo (posible causa de descuadre)
        ingresosSinFecha: recon.ingresosSinFecha,
        ingresosSinFechaMonto: recon.ingresosSinFechaMonto
      }
    };
  } catch(ex){ return { ok:false, error:ex.message+' (L:'+ex.lineNumber+')' }; }
}
