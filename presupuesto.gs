/* ==============================================================
   presupuesto.gs — Motor de Presupuesto y Proyección
   --------------------------------------------------------------
   Proyecta el siguiente trimestre con el modelo "estacional con
   piso creciente" (ratchet):

     Piso        = MAX( mismo trimestre del año anterior ,
                        último trimestre completo )
     Crecimiento = MAX( crecimiento real interanual ,
                        % objetivo de la hoja )
     Proyección  = MAX( base_estacional × (1 + crecimiento) , piso )

   Garantiza que la proyección nunca quede por debajo del año
   anterior NI por debajo del nivel reciente ya alcanzado.

   Metas: pestaña "Presupuesto_Metas" en ER_SS_ID.
   Histórico de ventas: BD_Ingresos (INGRESOS_SS_ID).
   Histórico de egresos: EGRESOS_IDS por año.

   Rutas:
     GET  ?action=presupuesto
     POST {action:'setupPresupuesto'}
     POST {action:'savePresupuestoMeta', periodo, linea, metaIngreso, crecimiento, ...}
   ============================================================== */

// Techo de crecimiento: la clínica arrancó a mediados de 2024 con ventas
// muy bajas, así que el crecimiento interanual crudo se dispara (>200%).
// Este tope mantiene la proyección realista (el piso ya cuida el mínimo).
var PRES_CREC_MAX = 0.30;   // 30% — ajustable
var PRES_METAS_TAB = 'Presupuesto_Metas';
var PRES_METAS_HEADERS = ['Periodo', 'Línea de servicio', 'Meta ingresos', 'Meta margen %', 'Crecimiento objetivo %', 'Notas'];

function _presNum(v) { if (typeof v === 'number') return v; var n = parseFloat(String(v || '').replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
function _presQ(month) { return Math.floor((month - 1) / 3) + 1; }          // mes 1-12 → trimestre 1-4
function _presQKey(y, q) { return y + '-Q' + q; }
function _presPrevQ(y, q) { return q > 1 ? { y: y, q: q - 1 } : { y: y - 1, q: 4 }; }
function _presNextQ(y, q) { return q < 4 ? { y: y, q: q + 1 } : { y: y + 1, q: 1 }; }
/* ¿El trimestre ('YYYY-Qn') ya terminó? (su último día es anterior a hoy) */
// Periodos abiertos manualmente (candado): override del bloqueo por "trimestre terminado".
function _presPeriodoAbierto(periodo){
  try{ var arr=JSON.parse(PropertiesService.getScriptProperties().getProperty('PRES_ABIERTOS')||'[]'); return arr.indexOf(String(periodo||'').trim())>-1; }catch(e){ return false; }
}
function _presQEnded(periodo) {
  var m = String(periodo || '').trim().match(/^(\d{4})-Q([1-4])$/); if (!m) return false;
  if (_presPeriodoAbierto(periodo)) return false;   // periodo abierto → editable aunque haya terminado
  var y = parseInt(m[1], 10), q = parseInt(m[2], 10);
  var fin = new Date(y, q * 3, 0);          // último día del último mes del trimestre
  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  return fin < hoy;
}
// Abre/cierra un trimestre para edición (candado). Gate de permisos en el frontend.
function presSetPeriodoAbierto(body){
  try{
    var periodo=String(body.periodo||'').trim();
    if(!/^\d{4}-Q[1-4]$/.test(periodo)) return {ok:false, error:'Periodo inválido (usa YYYY-Qn)'};
    var props=PropertiesService.getScriptProperties();
    var arr=JSON.parse(props.getProperty('PRES_ABIERTOS')||'[]');
    var i=arr.indexOf(periodo);
    if(body.abierto){ if(i<0) arr.push(periodo); } else { if(i>-1) arr.splice(i,1); }
    props.setProperty('PRES_ABIERTOS', JSON.stringify(arr));
    try{ logAudit(body.usuario||'sistema','Presupuesto', body.abierto?'Abrir periodo':'Cerrar periodo', periodo,'','',''); }catch(e){}
    return {ok:true, periodo:periodo, abierto:!!body.abierto, abiertos:arr};
  }catch(ex){ return {ok:false, error:ex.message}; }
}

/* ── Diagnóstico: vuelca la estructura real de la pestaña Budget ──
   Correr desde el editor y revisar "Registro de ejecución" (Ctrl+Enter).
   Sirve para mapear qué columna tiene el presupuesto por periodo/línea. */
function peekBudget() {
  var ss = SpreadsheetApp.openById(ER_SS_ID);
  var sh = ss.getSheetByName('Budget');
  if (!sh) { Logger.log('No existe la pestaña "Budget".'); return; }
  var nr = Math.min(sh.getLastRow(), 30), nc = Math.min(sh.getLastColumn(), 16);
  var vals = sh.getRange(1, 1, nr, nc).getValues();
  for (var i = 0; i < vals.length; i++) {
    Logger.log((i + 1) + ': ' + vals[i].map(function (v) {
      if (v instanceof Date) return v.getFullYear() + '-' + (v.getMonth() + 1);
      return String(v);
    }).join(' | '));
  }
}

/* ── Setup de la hoja de metas ──────────────────────────────────── */
function setupPresupuesto() {
  try {
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_METAS_TAB);
    if (!sh) sh = ss.insertSheet(PRES_METAS_TAB);
    sh.getRange(1, 1, 1, PRES_METAS_HEADERS.length).setValues([PRES_METAS_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 30);
    var widths = [110, 220, 140, 120, 160, 240];
    for (var w = 0; w < widths.length; w++) sh.setColumnWidth(w + 1, widths[w]);
    try {
      var bandings = sh.getBandings();
      for (var b = 0; b < bandings.length; b++) bandings[b].remove();
      sh.getRange(1, 1, sh.getMaxRows(), PRES_METAS_HEADERS.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    } catch (e) {}
    return { ok: true, message: 'Hoja "' + PRES_METAS_TAB + '" lista.', tab: PRES_METAS_TAB };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Lee histórico de ventas: por trimestre y por mes, por línea ── */
function _presHistoricoIngresos() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = null, all = ss.getSheets();
  for (var i = 0; i < all.length; i++) if (all[i].getName() === BD_INGRESOS_TAB) { sh = all[i]; break; }
  if (!sh) return { ok: false, error: 'No existe ' + BD_INGRESOS_TAB, q: {}, m: {}, lineas: [] };
  var raw = sh.getDataRange().getValues();
  var q = {}, m = {}, lineSet = {};
  for (var r = 1; r < raw.length; r++) {
    var row = raw[r];
    if (!String(row[0] || '').trim()) continue;     // sin OP
    var f = row[2];
    var d = (f instanceof Date) ? f : new Date(f);
    if (!d || isNaN(d.getTime())) continue;
    var y = d.getFullYear(), mo = d.getMonth() + 1, qq = _presQ(mo);
    var cat = String(row[4] || '').trim() || 'Sin categoría';
    var monto = _presNum(row[9]);                   // Total a pagar (venta)
    if (!monto) continue;
    lineSet[cat] = 1;
    var qk = _presQKey(y, qq);
    if (!q[qk]) q[qk] = { __total: 0 };
    q[qk][cat] = (q[qk][cat] || 0) + monto; q[qk].__total += monto;
    var mk = y + '-' + String(mo).padStart(2, '0');
    if (!m[mk]) m[mk] = { __total: 0 };
    m[mk][cat] = (m[mk][cat] || 0) + monto; m[mk].__total += monto;
  }
  return { ok: true, q: q, m: m, lineas: Object.keys(lineSet).sort() };
}

/* ── Histórico de egresos por trimestre, por subtipo + contable ─── */
function _presHistoricoEgresos() {
  var q = {}, m = {}, subSet = {}, contableBySub = {};
  Object.keys(EGRESOS_IDS).forEach(function (anio) {
    try {
      var ss = SpreadsheetApp.openById(EGRESOS_IDS[anio]);
      var tab = EGRESOS_TABS[anio];
      var sh = ss.getSheetByName(tab) || ss.getSheets()[0];
      if (!sh) return;
      var raw = sh.getDataRange().getValues();
      if (raw.length < 2) return;
      var H = raw[0].map(function (h) { return String(h).trim().toLowerCase(); });
      function col(kw) { for (var c = 0; c < H.length; c++) if (H[c].indexOf(kw) > -1) return c; return -1; }
      var iF = col('fecha'), iE = col('egresos'), iV = col('vencimiento'),
          iS = col('subtipo'), iC = col('contable');
      for (var r = 1; r < raw.length; r++) {
        var row = raw[r];
        var f = (iF > -1 ? row[iF] : '') || (iV > -1 ? row[iV] : '');
        var d = (f instanceof Date) ? f : new Date(f);
        if (!d || isNaN(d.getTime())) continue;
        var monto = _presNum(iE > -1 ? row[iE] : 0);
        if (!monto) continue;
        var qk = _presQKey(d.getFullYear(), _presQ(d.getMonth() + 1));
        var mk = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
        var sub = (iS > -1 ? String(row[iS] || '').trim() : '') || 'Otros';
        var cont = (iC > -1 ? String(row[iC] || '').trim() : '') || 'Gasto';
        if (!q[qk]) q[qk] = { __total: 0 };
        q[qk][sub] = (q[qk][sub] || 0) + monto; q[qk].__total += monto;
        if (!m[mk]) m[mk] = { __total: 0 };
        m[mk].__total += monto;
        subSet[sub] = 1;
        if (cont) contableBySub[sub] = cont;
      }
    } catch (e) {}
  });
  return { q: q, m: m, subtipos: Object.keys(subSet), contableBySub: contableBySub };
}

/* ══ AGRUPACIÓN DEL PRESUPUESTO (config) — grupos personalizados de productos
   para capturar/mover la meta por bucket, no producto por producto. Hoja
   'Presupuesto_Grupos' (Grupo | Producto | IncluirMaestro). ══════════════ */
var PRES_GRUPOS_TAB = 'Presupuesto_Grupos';
// Hoja: Grupo | Subgrupo | Producto | IncluirMaestro | Orden  (subgrupo/producto pueden ir vacíos)
function readGruposPresupuesto() {
  try {
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_GRUPOS_TAB);
    if (!sh) return { ok: true, grupos: [], productoARuta: {}, _setup: false };
    var raw = sh.getDataRange().getValues();
    var gInfo = {}, order = [], productoARuta = {};
    for (var i = 1; i < raw.length; i++) {
      var g = String(raw[i][0] || '').trim(); if (!g) continue;
      var sub = String(raw[i][1] || '').trim();
      var p = String(raw[i][2] || '').trim();
      var inclRaw = raw[i][3];
      var incl = inclRaw === true || String(inclRaw).toUpperCase() === 'TRUE' || inclRaw === '' || inclRaw == null;
      var ord = Number(raw[i][4]) || 0;
      if (!gInfo[g]) { gInfo[g] = { nombre: g, incluirMaestro: incl, orden: ord, subs: [], _subset: {} }; order.push(g); }
      gInfo[g].incluirMaestro = incl; gInfo[g].orden = ord;
      if (sub && !gInfo[g]._subset[sub]) { gInfo[g]._subset[sub] = 1; gInfo[g].subs.push(sub); }
      if (p) productoARuta[p] = { grupo: g, sub: sub };
    }
    order.sort(function (a, b) { return (gInfo[a].orden - gInfo[b].orden) || a.localeCompare(b); });
    var grupos = order.map(function (g) { var G = gInfo[g]; return { nombre: G.nombre, incluirMaestro: G.incluirMaestro, orden: G.orden, subs: G.subs }; });
    return { ok: true, grupos: grupos, productoARuta: productoARuta, _setup: true };
  } catch (ex) { return { ok: false, error: ex.message, grupos: [], productoARuta: {} }; }
}
// body.grupos: [{nombre, incluirMaestro, orden, directos:[prod...], subgrupos:[{nombre, productos:[prod...]}]}]
function saveGruposPresupuesto(body) {
  try {
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_GRUPOS_TAB);
    if (!sh) sh = ss.insertSheet(PRES_GRUPOS_TAB);
    sh.clear();
    sh.getRange(1, 1, 1, 5).setValues([['Grupo', 'Subgrupo', 'Producto', 'IncluirMaestro', 'Orden']]);
    var rows = [];
    (body.grupos || []).forEach(function (g, gi) {
      var nom = String(g.nombre || '').trim(); if (!nom) return;
      var incl = g.incluirMaestro !== false;
      var ord = (g.orden != null) ? g.orden : gi;
      var wrote = false;
      (g.directos || g.productos || []).forEach(function (p) { p = String(p || '').trim(); if (p) { rows.push([nom, '', p, incl, ord]); wrote = true; } });
      (g.subgrupos || []).forEach(function (s) {
        var sn = String(s.nombre || '').trim(); if (!sn) return;
        var prods = s.productos || [];
        if (!prods.length) { rows.push([nom, sn, '', incl, ord]); wrote = true; }
        prods.forEach(function (p) { p = String(p || '').trim(); if (p) { rows.push([nom, sn, p, incl, ord]); wrote = true; } });
      });
      if (!wrote) rows.push([nom, '', '', incl, ord]);
    });
    if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
    try { logAudit(body.usuario || 'sistema', 'Presupuesto', 'Guardar grupos', '', '', '', rows.length + ' filas'); } catch (e) {}
    return { ok: true, guardados: rows.length };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ══ Candados "movible" por grupo / subgrupo(ciclo) ═══════════════════════
   Persisten qué buckets se mueven con % maestro / cuadrar / recomendado y
   cuáles quedan FIJOS al número capturado. Script Property PRES_MAESTRO_LOCKS
   = { "Grupo": false, "Grupo|Ciclo": false }. Ausente = movible (default). */
function readPresLocks() {
  try { var raw = PropertiesService.getScriptProperties().getProperty('PRES_MAESTRO_LOCKS'); return { ok: true, locks: raw ? JSON.parse(raw) : {} }; }
  catch (e) { return { ok: true, locks: {} }; }
}
function presSetLock(body) {
  try {
    var k = String(body.key || '').trim(); if (!k) return { ok: false, error: 'sin key' };
    var p = PropertiesService.getScriptProperties();
    var raw = p.getProperty('PRES_MAESTRO_LOCKS'); var m = raw ? JSON.parse(raw) : {};
    if (body.movible === false) m[k] = false; else delete m[k]; // movible=true → quitar candado
    p.setProperty('PRES_MAESTRO_LOCKS', JSON.stringify(m));
    return { ok: true, locks: m };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ══ % de los escenarios Conservador / Optimista (sobre la Base) ═══════════
   El usuario los define porque el mercado varía; NO están fijos en el código.
   Script Property PRES_ESCENARIOS = { conservador: 10, optimista: 8 } (en %). */
function _presEscenariosCfg() {
  try { var raw = PropertiesService.getScriptProperties().getProperty('PRES_ESCENARIOS'); var m = raw ? JSON.parse(raw) : {};
    return { conservador: (m.conservador != null ? Number(m.conservador) : 10), optimista: (m.optimista != null ? Number(m.optimista) : 8) };
  } catch (e) { return { conservador: 10, optimista: 8 }; }
}
function presSetEscenarios(body) {
  try {
    var p = PropertiesService.getScriptProperties();
    var raw = p.getProperty('PRES_ESCENARIOS'); var m = raw ? JSON.parse(raw) : {};
    if (body.conservador !== undefined && body.conservador !== '' && body.conservador !== null) m.conservador = Math.max(0, Number(body.conservador) || 0);
    if (body.optimista   !== undefined && body.optimista   !== '' && body.optimista   !== null) m.optimista   = Math.max(0, Number(body.optimista)   || 0);
    p.setProperty('PRES_ESCENARIOS', JSON.stringify(m));
    return { ok: true, escenarios: m };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Histórico de metas por trimestre (últimos ~7 Q + el siguiente proyectado):
   meta fijada vs real logrado + cumplimiento, para ver la racha. Solo 2 lecturas. */
function readHistoricoMetas() {
  try {
    var hist = _presHistoricoIngresos();
    var egHist = _presHistoricoEgresos();       // egresos reales por trimestre (para ganancia/pérdida)
    var metasInfo = _presLeerMetas();
    var histQ = (hist && hist.q) || {};
    var egQ = (egHist && egHist.q) || {};
    var metas = (metasInfo && metasInfo.map) || {};
    var hoy = new Date(), y = hoy.getFullYear(), q = _presQ(hoy.getMonth() + 1);
    var list = [], yy = y, qq = q;
    for (var i = 0; i < 7; i++) { list.unshift({ y: yy, q: qq }); qq--; if (qq < 1) { qq = 4; yy--; } }
    var yn = y, qn = q + 1; if (qn > 4) { qn = 1; yn++; }
    list.push({ y: yn, q: qn });
    var curKey = _presQKey(y, q), nextKey = _presQKey(yn, qn);
    var out = list.map(function (p) {
      var per = _presQKey(p.y, p.q);
      var meta = (metas[per] && metas[per].__total) || 0;
      var real = (histQ[per] && histQ[per].__total) || 0;
      var egresoReal = (egQ[per] && egQ[per].__total) || 0;
      var tipo = per === nextKey ? 'proyeccion' : (per === curKey ? 'actual' : 'pasado');
      return { periodo: per, tipo: tipo, meta: meta, real: real,
        cumplimiento: meta > 0 ? Math.round(real / meta * 100) : null,
        alcanzada: meta > 0 ? (real >= meta) : null,
        egresoReal: egresoReal,
        gananciaReal: real - egresoReal };          // ganancia (o pérdida si es negativo) real del trimestre
    });
    return { ok: true, historico: out, trimestreActual: curKey };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Lee metas del almacén que administra la PÁGINA (Presupuesto_Metas) ─
   El usuario nunca edita esta hoja a mano: la página escribe vía
   savePresupuestoMeta. Una fila con Línea = "TOTAL" fija el total del
   trimestre; filas por línea suman al total si no hay TOTAL.            */
function _presLeerMetas() {
  var map = {};
  try {
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_METAS_TAB);
    if (!sh) return { map: map, _setup: false };
    var raw = sh.getDataRange().getValues();
    for (var r = 1; r < raw.length; r++) {
      var per = String(raw[r][0] || '').trim();
      var lin = String(raw[r][1] || '').trim();
      if (!per || !lin) continue;
      var meta = _presNum(raw[r][2]);
      if (!map[per]) map[per] = { __total: 0, __totalEg: 0, __tot: null, _lineas: {}, _egLineas: {} };
      if (lin.toUpperCase() === 'TOTAL') {
        map[per].__tot = meta;
      } else if (lin.substring(0, 3).toUpperCase() === 'EG:') {
        // Meta de EGRESO (prefijo EG:)
        map[per]._egLineas[lin.substring(3)] = meta;
        map[per].__totalEg += meta;
      } else {
        map[per]._lineas[lin] = { metaIngreso: meta, metaMargen: _presNum(raw[r][3]), crecObjetivo: _presNum(raw[r][4]) / 100 };
        map[per].__total += meta;
      }
    }
    // El TOTAL explícito ("Fijar meta del trimestre") MANDA sobre la suma por línea.
    // Si no hay fila TOTAL, el total es la suma de las metas por producto/línea.
    Object.keys(map).forEach(function (p) {
      if (map[p].__tot != null) map[p].__total = map[p].__tot;
    });
    return { map: map, _setup: true };
  } catch (e) { return { map: map, _setup: false }; }
}

/* ── Crecimiento real interanual (últimos 4 trim. vs 4 previos) ── */
function _presCrecimientoReal(histQ, refY, refQ) {
  var ult = 0, prev = 0, yy = refY, qq = refQ;
  for (var i = 0; i < 4; i++) { var p = _presPrevQ(yy, qq); var k = _presQKey(p.y, p.q); ult += (histQ[k] && histQ[k].__total) || 0; yy = p.y; qq = p.q; }
  for (var j = 0; j < 4; j++) { var p2 = _presPrevQ(yy, qq); var k2 = _presQKey(p2.y, p2.q); prev += (histQ[k2] && histQ[k2].__total) || 0; yy = p2.y; qq = p2.q; }
  if (prev <= 0) return 0;
  return (ult - prev) / prev;
}

/* ── Motor principal ────────────────────────────────────────────── */
/* Crecimiento interanual de una serie total por trimestre (trailing 4Q vs prev 4Q). */
function _presCrecCantTotal(map, refY, refQ) {
  var ult = 0, prev = 0, yy = refY, qq = refQ;
  for (var i = 0; i < 4; i++) { var p = _presPrevQ(yy, qq); ult += map[_presQKey(p.y, p.q)] || 0; yy = p.y; qq = p.q; }
  for (var j = 0; j < 4; j++) { var p2 = _presPrevQ(yy, qq); prev += map[_presQKey(p2.y, p2.q)] || 0; yy = p2.y; qq = p2.q; }
  if (prev <= 0) return null;
  return (ult - prev) / prev;
}

/* ── Proyección de INGRESOS por producto = cantidad × precio promedio real,
   agrupada como el Board Deck (columna U → categoría → producto).
   Reutiliza helpers de summary.gs (_summaryReadIngresos, _summaryRevOrden,
   _sumOrdIn, SUMMARY_SUBGROUP_ORDER, SUMMARY_PRODUCT_ORDER). ───────────── */
function _presIngresosProy(tY, tQ) {
  var q = {}, qCantTot = {};
  var thisYear = new Date().getFullYear();
  for (var y = thisYear - 2; y <= thisYear; y++) {
    var rows; try { rows = _summaryReadIngresos(y); } catch (e) { rows = []; }
    rows.forEach(function (r) {
      var f = (r.fecha || '').substring(0, 10); if (f.length < 7) return;
      var yy = parseInt(f.substring(0, 4), 10), mo = parseInt(f.substring(5, 7), 10);
      var qk = _presQKey(yy, _presQ(mo));
      var g = String(r.grupoU || '').trim() || String(r.categoria || '').trim() || '(Sin grupo)';
      var s = String(r.categoria || '').trim() || g;
      var p = String(r.producto || '').trim() || s;
      var cant = Number(r.cantidad) || 0, imp = Number(r.total) || 0;
      if (!q[qk]) q[qk] = {}; if (!q[qk][g]) q[qk][g] = {}; if (!q[qk][g][s]) q[qk][g][s] = {};
      if (!q[qk][g][s][p]) q[qk][g][s][p] = { cant: 0, imp: 0 };
      q[qk][g][s][p].cant += cant; q[qk][g][s][p].imp += imp;
      qCantTot[qk] = (qCantTot[qk] || 0) + cant;
    });
  }
  var hoy = new Date(), cy = hoy.getFullYear(), cq = _presQ(hoy.getMonth() + 1);
  var rec = _presPrevQ(cy, cq), kRec = _presQKey(rec.y, rec.q), kAA = _presQKey(tY - 1, tQ);
  var gTot = _presCrecCantTotal(qCantTot, cy, cq); if (gTot === null) gTot = 0;
  gTot = Math.min(Math.max(gTot, 0), PRES_CREC_MAX);

  function cellP(k, g, s, p) { return (((q[k] || {})[g] || {})[s] || {})[p] || { cant: 0, imp: 0 }; }
  var allG = {};
  [kRec, kAA].forEach(function (k) { Object.keys(q[k] || {}).forEach(function (g) { allG[g] = 1; }); });
  var grupos = [], totImp = 0;
  Object.keys(allG).forEach(function (g) {
    var subMap = {};
    [kRec, kAA].forEach(function (k) { Object.keys((q[k] || {})[g] || {}).forEach(function (s) { subMap[s] = 1; }); });
    var subgrupos = [], gImp = 0, gCant = 0, gAA = 0, gRe = 0;
    Object.keys(subMap).forEach(function (s) {
      var prodMap = {};
      [kRec, kAA].forEach(function (k) { Object.keys(((q[k] || {})[g] || {})[s] || {}).forEach(function (p) { prodMap[p] = 1; }); });
      var productos = [], sImp = 0, sCant = 0;
      Object.keys(prodMap).forEach(function (p) {
        var re = cellP(kRec, g, s, p), aa = cellP(kAA, g, s, p);
        var precio = re.cant > 0 ? re.imp / re.cant : (aa.cant > 0 ? aa.imp / aa.cant : 0);
        var baseCant = Math.max(aa.cant, re.cant);
        var cantProy = Math.round(baseCant * (1 + gTot)); if (cantProy < Math.round(baseCant)) cantProy = Math.round(baseCant);
        var impProy = cantProy * precio;
        if (impProy <= 0 && re.imp <= 0 && aa.imp <= 0) return;
        productos.push({ producto: p, cantAnioAnt: aa.cant, cantReciente: re.cant, cantProy: cantProy,
          precio: precio, importeProy: impProy, impReciente: re.imp, impAnioAnt: aa.imp });
        sImp += impProy; sCant += cantProy; gAA += aa.imp; gRe += re.imp;
      });
      if (!productos.length) return;
      productos.sort(function (a, b) { return (_sumOrdIn(SUMMARY_PRODUCT_ORDER, a.producto) - _sumOrdIn(SUMMARY_PRODUCT_ORDER, b.producto)) || (b.importeProy - a.importeProy); });
      subgrupos.push({ sub: s, importeProy: sImp, cantProy: sCant, productos: productos });
      gImp += sImp; gCant += sCant;
    });
    if (gImp <= 0.5) return;
    subgrupos.sort(function (a, b) { return (_sumOrdIn(SUMMARY_SUBGROUP_ORDER, a.sub) - _sumOrdIn(SUMMARY_SUBGROUP_ORDER, b.sub)) || (b.importeProy - a.importeProy); });
    grupos.push({ grupo: g, orden: _summaryRevOrden(g), importeProy: gImp, cantProy: gCant, anioAntImp: gAA, recienteImp: gRe, subgrupos: subgrupos });
    totImp += gImp;
  });
  grupos.sort(function (a, b) { return (a.orden - b.orden) || (b.importeProy - a.importeProy); });
  return { grupos: grupos, totalImporte: totImp, crecimientoCant: gTot, kReciente: kRec, kAnioAnt: kAA };
}

function readPresupuesto(periodo) {
  try {
    var hi = _presHistoricoIngresos();
    if (!hi.ok) return { ok: false, error: hi.error };
    var histQ = hi.q, histM = hi.m, lineas = hi.lineas;
    var metasInfo = _presLeerMetas();
    var metas = metasInfo.map;
    var egQ = _presHistoricoEgresos();

    var hoy = new Date();
    var curY = hoy.getFullYear(), curQ = _presQ(hoy.getMonth() + 1);
    var nx = _presNextQ(curY, curQ);
    var tgtY = nx.y, tgtQ = nx.q;
    // Permite consultar un trimestre específico (histórico o futuro): periodo = 'YYYY-Qn'
    var esHistorico = false;
    if (periodo && /^\d{4}-Q[1-4]$/.test(String(periodo).trim())) {
      var pm = String(periodo).trim().match(/^(\d{4})-Q([1-4])$/);
      tgtY = parseInt(pm[1], 10); tgtQ = parseInt(pm[2], 10);
      // histórico = SOLO si el trimestre YA TERMINÓ. El trimestre EN CURSO y los
      // futuros son editables (permite ajustar el Q actual apenas empezado).
      // Un periodo ABIERTO manualmente (candado) también se vuelve editable.
      esHistorico = ((tgtY < curY) || (tgtY === curY && tgtQ < curQ)) && !_presPeriodoAbierto(String(periodo).trim());
    }
    var perActual = _presQKey(curY, curQ);
    var perSig = _presQKey(tgtY, tgtQ);
    // Motor nuevo: ingresos por producto (cantidad × precio real), agrupado como Board Deck
    var ingProd = _presIngresosProy(tgtY, tgtQ);

    // Crecimiento real interanual (total) — referencia para líneas sin histórico propio
    var gRealTotal = _presCrecimientoReal(histQ, curY, curQ);

    // Último trimestre completo (piso reciente) = el inmediato anterior al actual
    var lastComplete = _presPrevQ(curY, curQ);
    var kReciente = _presQKey(lastComplete.y, lastComplete.q);
    var kAnioAnt = _presQKey(tgtY - 1, tgtQ);

    var metaSig = (metas[perSig] && metas[perSig]._lineas) || {};

    // Proyección por línea
    var lineasProy = [], totBase = 0, totProy = 0, totCons = 0, totOpt = 0, totAnioAnt = 0, totReciente = 0, totMeta = 0;
    lineas.forEach(function (cat) {
      var anioAnt = (histQ[kAnioAnt] && histQ[kAnioAnt][cat]) || 0;
      var reciente = (histQ[kReciente] && histQ[kReciente][cat]) || 0;
      // crecimiento real de la línea (su YoY); fallback al total
      var gLin = _presCrecimientoLinea(histQ, cat, curY, curQ);
      if (gLin === null) gLin = gRealTotal;
      var gObj = (metaSig[cat] && metaSig[cat].crecObjetivo) || 0;
      // crecimiento = el mayor entre real y objetivo, pero con techo (evita
      // que el arranque 2024 dispare la proyección) y nunca negativo.
      var g = Math.min(Math.max(gLin, gObj, 0), PRES_CREC_MAX);
      var baseEstacional = anioAnt > 0 ? anioAnt : reciente; // si no hay año anterior, usa reciente
      var proyEstacional = baseEstacional * (1 + g);
      var piso = Math.max(anioAnt, reciente);
      var proy = Math.max(proyEstacional, piso);
      var conservador = Math.max(piso, baseEstacional * (1 + Math.max(g * 0.5, 0.02)));
      var optimista = proy * 1.12;
      var meta = (metaSig[cat] && metaSig[cat].metaIngreso) || 0;
      lineasProy.push({
        linea: cat, anioAnterior: anioAnt, reciente: reciente, piso: piso,
        crecimientoUsado: g, base: baseEstacional,
        conservador: conservador, proyeccion: proy, optimista: optimista, meta: meta
      });
      totBase += baseEstacional; totProy += proy; totCons += conservador; totOpt += optimista;
      totAnioAnt += anioAnt; totReciente += reciente; totMeta += meta;
    });
    lineasProy.sort(function (a, b) { return b.proyeccion - a.proyeccion; });

    // ── Trimestre en curso: pace ──
    var qStart = new Date(curY, (curQ - 1) * 3, 1);
    var qEnd = new Date(curY, curQ * 3, 0);
    var diasTotales = Math.round((qEnd - qStart) / 86400000) + 1;
    var diasTransc = Math.min(Math.max(Math.round((hoy - qStart) / 86400000) + 1, 1), diasTotales);
    var realActual = (histQ[perActual] && histQ[perActual].__total) || 0;
    var fraccion = diasTransc / diasTotales;
    var proyCierre = fraccion > 0 ? realActual / fraccion : realActual;
    var metaActualTotal = (metas[perActual] && metas[perActual].__total) || 0;
    var cumplimiento = metaActualTotal > 0 ? proyCierre / metaActualTotal : null;
    var semaforo = cumplimiento === null ? 'sin-meta' : (cumplimiento >= 1 ? 'verde' : (cumplimiento >= 0.9 ? 'amarillo' : 'rojo'));

    // ── Egresos POR LÍNEA (subtipo) con el mismo modelo ratchet ──
    var egHistQ = egQ.q;
    var egMetaSig = (metas[perSig] && metas[perSig]._egLineas) || {};
    var gEgTotal = _presCrecimientoLinea(egHistQ, '__total', curY, curQ); if (gEgTotal === null) gEgTotal = 0;
    var egLineasProy = [], egTotAnioAnt = 0, egTotReciente = 0, egTotProy = 0, egTotMeta = 0;
    (egQ.subtipos || []).forEach(function (sub) {
      var aa = (egHistQ[kAnioAnt] && egHistQ[kAnioAnt][sub]) || 0;
      var rec = (egHistQ[kReciente] && egHistQ[kReciente][sub]) || 0;
      var gL = _presCrecimientoLinea(egHistQ, sub, curY, curQ); if (gL === null) gL = gEgTotal;
      var g = Math.min(Math.max(gL, 0), PRES_CREC_MAX);
      var base = aa > 0 ? aa : rec;
      var proy = Math.max(base * (1 + g), Math.max(aa, rec));
      if (proy <= 0 && rec <= 0 && aa <= 0) return;
      var meta = egMetaSig[sub] || 0;
      // Sección (COGS/OpEx/G&A/Taxes) y subgrupo (Payroll, Rent…) como el Board Deck
      var cont = egQ.contableBySub[sub] || 'Gasto';
      var seccion = 'GA', subg = sub;
      try { var c = _summaryDefaultClass('egreso', cont + '|' + sub); subg = _summaryEgSubgroup(sub, c.grupo);
        seccion = SUMMARY_EG_SUBGROUP_SECTION[_sumNorm(subg)] || c.grupo; } catch (e) {}
      egLineasProy.push({ linea: sub, grupo: cont, seccion: seccion, subgrupo: subg, anioAnterior: aa, reciente: rec, crecimientoUsado: g, proyeccion: proy, meta: meta });
      egTotAnioAnt += aa; egTotReciente += rec; egTotProy += proy; egTotMeta += meta;
    });
    egLineasProy.sort(function (a, b) { return b.proyeccion - a.proyeccion; });
    // Agrupar egresos: sección → subgrupo → subtipos (para render tipo Board Deck)
    var egSecOrden = { COGS: 1, OPEX: 2, GA: 3, TAXES: 4 };
    var egSecMap = {};
    egLineasProy.forEach(function (l) {
      var sec = l.seccion || 'GA';
      if (!egSecMap[sec]) egSecMap[sec] = { seccion: sec, orden: egSecOrden[sec] || 9, importeProy: 0, subs: {} };
      var S = egSecMap[sec]; S.importeProy += l.proyeccion;
      if (!S.subs[l.subgrupo]) S.subs[l.subgrupo] = { sub: l.subgrupo, importeProy: 0, lineas: [] };
      S.subs[l.subgrupo].importeProy += l.proyeccion; S.subs[l.subgrupo].lineas.push(l);
    });
    var egGrupos = Object.keys(egSecMap).map(function (k) {
      var S = egSecMap[k];
      S.subgrupos = Object.keys(S.subs).map(function (sk) { var sg = S.subs[sk]; sg.lineas.sort(function (a, b) { return b.proyeccion - a.proyeccion; }); return sg; })
        .sort(function (a, b) { return (_summaryEgSubgroupOrden(a.sub) - _summaryEgSubgroupOrden(b.sub)) || (b.importeProy - a.importeProy); });
      delete S.subs; return S;
    }).sort(function (a, b) { return a.orden - b.orden; });
    var egProy = egTotProy;
    // Ingresos proyectados = la MISMA cifra que muestra la tarjeta (por producto: cantidad × precio),
    // no la proyección vieja por línea. Antes el margen usaba totProy y salía inconsistente (negativo).
    var _ingProyTot = ingProd.totalImporte || totProy;
    var margenProy = _ingProyTot - egProy;
    var margenPct = _ingProyTot > 0 ? (margenProy / _ingProyTot) * 100 : 0;

    // Adjuntar budget guardado (meta por producto) a la estructura nueva de ingresos
    (ingProd.grupos || []).forEach(function (G) { (G.subgrupos || []).forEach(function (S) { (S.productos || []).forEach(function (p) {
      p.budget = (metaSig[p.producto] && metaSig[p.producto].metaIngreso) || 0;
    }); }); });

    // ── Tendencia mensual (income + egresos) para la gráfica ──
    var tendencia = _presTendencia(histM, egQ.m || {}, tgtY, tgtQ, _ingProyTot, egProy);

    return {
      ok: true,
      _setup: metasInfo._setup,
      hoy: hoy.getFullYear() + '-' + String(hoy.getMonth() + 1).padStart(2, '0') + '-' + String(hoy.getDate()).padStart(2, '0'),
      trimestreActual: perActual,
      trimestreSiguiente: perSig,
      crecimientoRealTotal: gRealTotal,
      crecimientoAplicado: Math.min(Math.max(gRealTotal, 0), PRES_CREC_MAX),
      crecimientoTope: PRES_CREC_MAX,
      actual: {
        periodo: perActual, meta: metaActualTotal, realAcumulado: realActual,
        diasTranscurridos: diasTransc, diasTotales: diasTotales,
        proyeccionCierre: proyCierre, cumplimientoPct: cumplimiento, semaforo: semaforo
      },
      periodoConsultado: perSig,
      esHistorico: esHistorico,
      periodoAbierto: _presPeriodoAbierto(perSig),   // candado: abierto manualmente para editar
      qTerminado: (function(){ var mm=String(perSig).match(/^(\d{4})-Q([1-4])$/); if(!mm) return false; return new Date(parseInt(mm[1]),parseInt(mm[2])*3,0) < new Date(new Date().setHours(0,0,0,0)); })(),
      realTrimestre: (histQ[perSig] && histQ[perSig].__total) || 0,   // lo que REALMENTE pasó ese Q (para histórico)
      egRealTrimestre: (egHistQ[perSig] && egHistQ[perSig].__total) || 0,
      siguiente: {
        periodo: perSig,
        lineas: lineasProy,
        ingresosGrupos: ingProd.grupos,              // NUEVO: cantidad × precio, agrupado como Board Deck
        ingresosTotalProy: ingProd.totalImporte,
        crecimientoCant: ingProd.crecimientoCant,
        totales: (function(){ var _b = ingProd.totalImporte || totProy; var _e = _presEscenariosCfg();
          return { anioAnterior: totAnioAnt, reciente: totReciente, base: totBase,
            conservador: Math.round(_b * (1 - (_e.conservador||0)/100)),
            proyeccion: _b,
            optimista: Math.round(_b * (1 + (_e.optimista||0)/100)),
            meta: (metas[perSig] && metas[perSig].__total) || 0 }; })()
      },
      escenariosPct: _presEscenariosCfg(),
      egresos: {
        reciente: egTotReciente, proyeccion: egProy,
        margenProyectado: margenProy, margenPct: margenPct,
        lineas: egLineasProy,
        grupos: egGrupos,                            // NUEVO: sección → subgrupo → subtipos
        totales: { anioAnterior: egTotAnioAnt, reciente: egTotReciente, proyeccion: egTotProy, meta: egTotMeta }
      },
      tendencia: tendencia,
      lineasDisponibles: lineas
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function _presCrecimientoLinea(histQ, cat, refY, refQ) {
  var ult = 0, prev = 0, yy = refY, qq = refQ;
  for (var i = 0; i < 4; i++) { var p = _presPrevQ(yy, qq); var k = _presQKey(p.y, p.q); ult += (histQ[k] && histQ[k][cat]) || 0; yy = p.y; qq = p.q; }
  for (var j = 0; j < 4; j++) { var p2 = _presPrevQ(yy, qq); var k2 = _presQKey(p2.y, p2.q); prev += (histQ[k2] && histQ[k2][cat]) || 0; yy = p2.y; qq = p2.q; }
  if (prev <= 0) return null;
  return (ult - prev) / prev;
}

/* ── Serie mensual para la gráfica: income + egresos ────────────── */
function _presTendencia(histM, egM, tgtY, tgtQ, totProySig, egProySig) {
  var MESES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  var meses = [], real = [], proy = [], egReal = [], egProy = [];
  var hoy = new Date();
  var y = hoy.getFullYear(), mo = hoy.getMonth() + 1;
  var startY = y, startM = mo - 9;
  while (startM <= 0) { startM += 12; startY--; }
  var yy = startY, mm = startM;
  for (var i = 0; i < 9; i++) {
    var mk = yy + '-' + String(mm).padStart(2, '0');
    meses.push(MESES[mm - 1] + ' ' + String(yy).slice(2));
    real.push((histM[mk] && histM[mk].__total) || 0);
    egReal.push((egM[mk] && egM[mk].__total) || 0);
    proy.push(null); egProy.push(null);
    mm++; if (mm > 12) { mm = 1; yy++; }
  }
  var qMonths = [(tgtQ - 1) * 3 + 1, (tgtQ - 1) * 3 + 2, (tgtQ - 1) * 3 + 3];
  function shape(map) { return qMonths.map(function (m2) { var k = (tgtY - 1) + '-' + String(m2).padStart(2, '0'); return (map[k] && map[k].__total) || 0; }); }
  var shInc = shape(histM), shEg = shape(egM);
  var sumInc = shInc.reduce(function (a, b) { return a + b; }, 0), sumEg = shEg.reduce(function (a, b) { return a + b; }, 0);
  // conectar las proyecciones desde el último real
  proy[proy.length - 1] = real[real.length - 1] || 0;
  egProy[egProy.length - 1] = egReal[egReal.length - 1] || 0;
  qMonths.forEach(function (m2, idx) {
    meses.push(MESES[m2 - 1] + ' ' + String(tgtY).slice(2));
    real.push(null); egReal.push(null);
    proy.push(totProySig * (sumInc > 0 ? shInc[idx] / sumInc : 1 / 3));
    egProy.push(egProySig * (sumEg > 0 ? shEg[idx] / sumEg : 1 / 3));
  });
  return { meses: meses, real: real, proyeccion: proy, egresoReal: egReal, egresoProy: egProy };
}

/* ── Guardar / actualizar una meta ──────────────────────────────── */
// Borra la(s) fila(s) TOTAL de un periodo (para que la suma por producto vuelva a gobernar).
function _presClearTotal(sh, per) {
  var lr = sh.getLastRow(); if (lr < 2) return;
  var vals = sh.getRange(2, 1, lr - 1, 2).getValues();
  for (var i = vals.length - 1; i >= 0; i--) { // de abajo hacia arriba para no romper índices
    if (String(vals[i][0]).trim() === per && String(vals[i][1]).trim().toUpperCase() === 'TOTAL') sh.deleteRow(i + 2);
  }
}
function savePresupuestoMeta(body) {
  try {
    if (!body || !String(body.periodo || '').trim() || !String(body.linea || '').trim())
      return { ok: false, error: 'Periodo y línea son obligatorios.' };
    if (_presQEnded(body.periodo)) return { ok: false, error: 'Ese trimestre ya terminó; el presupuesto no se puede modificar.' };
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_METAS_TAB);
    if (!sh) { setupPresupuesto(); sh = ss.getSheetByName(PRES_METAS_TAB); }
    var per = String(body.periodo).trim(), lin = String(body.linea).trim();
    // Si capturan una meta POR PRODUCTO, se retira el TOTAL fijo para que gobierne la suma (última acción manda).
    if (lin.toUpperCase() !== 'TOTAL') { try { _presClearTotal(sh, per); } catch (e) {} }
    var fila = [per, lin, _presNum(body.metaIngreso), _presNum(body.metaMargen), _presNum(body.crecimiento), String(body.notas || '')];
    // Buscar fila existente (mismo periodo + línea)
    var lr = sh.getLastRow(), found = 0;
    if (lr > 1) {
      var vals = sh.getRange(2, 1, lr - 1, 2).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim() === per && String(vals[i][1]).trim() === lin) { found = i + 2; break; }
      }
    }
    if (found) sh.getRange(found, 1, 1, fila.length).setValues([fila]);
    else sh.appendRow(fila);
    try { logAudit(body.usuario || '', 'Presupuesto', found ? 'Edición meta' : 'Alta meta', per + ' · ' + lin, 'meta', '', body.metaIngreso); } catch (e) {}
    return { ok: true, message: 'Meta guardada.' };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Guarda varias metas por línea de un periodo (una sola llamada) ── */
function savePresupuestoMetasBatch(body) {
  try {
    if (!body || !String(body.periodo || '').trim() || !Array.isArray(body.metas))
      return { ok: false, error: 'Datos incompletos.' };
    if (_presQEnded(body.periodo)) return { ok: false, error: 'Ese trimestre ya terminó; el presupuesto no se puede modificar.' };
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_METAS_TAB);
    if (!sh) { setupPresupuesto(); sh = ss.getSheetByName(PRES_METAS_TAB); }
    var per = String(body.periodo).trim();
    try { _presClearTotal(sh, per); } catch (e) {} // captura por producto → quita el TOTAL fijo
    var lr = sh.getLastRow(), existing = {};
    if (lr > 1) {
      var vals = sh.getRange(2, 1, lr - 1, 2).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim() === per) existing[String(vals[i][1]).trim().toLowerCase()] = i + 2;
      }
    }
    var appends = [];
    body.metas.forEach(function (m) {
      var lin = String(m.linea || '').trim(); if (!lin) return;
      var row = [per, lin, _presNum(m.metaIngreso), 0, 0, ''];
      var found = existing[lin.toLowerCase()];
      if (found) sh.getRange(found, 1, 1, row.length).setValues([row]);
      else appends.push(row);
    });
    if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, appends[0].length).setValues(appends);
    return { ok: true, message: body.metas.length + ' metas guardadas.' };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
