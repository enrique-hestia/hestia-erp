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
// Col 7 "Meta cantidad": la CANTIDAD manual capturada por producto. El grid de ingresos
// es cantidad×precio; guardar solo el importe obligaba a reconstruir la cantidad dividiendo
// por un precio que se recalcula en vivo (deriva). Persistir la cantidad la vuelve exacta.
var PRES_METAS_HEADERS = ['Periodo', 'Línea de servicio', 'Meta ingresos', 'Meta margen %', 'Crecimiento objetivo %', 'Notas', 'Meta cantidad'];

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
    var widths = [110, 220, 140, 120, 160, 240, 120];
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

/* ══ PARÁMETROS DE PROYECCIÓN (escenarios) POR PERIODO ════════════════════
   Persisten los %s del panel de escenarios y el escenario activo por trimestre,
   en la MISMA hoja Presupuesto_Metas, como una fila especial:
     Periodo | '__PROYECCION__' | Conservador % | Optimista % | Base % (vs recom) | Notas=escenario
   El recomendado del modelo es solo la SEMILLA: una vez que el usuario guarda
   sus %s, esos MANDAN y no se sobrescriben al recargar. Guardado PARCIAL: cada
   campo omitido conserva su valor previo (o la semilla si nunca se guardó).     */
function savePresupuestoProyeccion(body) {
  try {
    var per = String((body && body.periodo) || '').trim();
    if (!/^\d{4}-Q[1-4]$/.test(per)) return { ok: false, error: 'Periodo inválido (usa YYYY-Qn).' };
    if (_presQEnded(per)) return { ok: false, error: 'Ese trimestre ya terminó; el presupuesto no se puede modificar.' };
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_METAS_TAB);
    if (!sh) { setupPresupuesto(); sh = ss.getSheetByName(PRES_METAS_TAB); }
    // Valores previos (o la semilla global) para permitir guardado parcial.
    var seed = _presEscenariosCfg();
    var prev = { conservador: seed.conservador, optimista: seed.optimista, basePct: 0, seleccionado: 'base' };
    var lr = sh.getLastRow(), found = 0;
    if (lr > 1) {
      var vals = sh.getRange(2, 1, lr - 1, 6).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim() === per && String(vals[i][1]).trim().toUpperCase() === '__PROYECCION__') {
          found = i + 2;
          prev = { conservador: _presNum(vals[i][2]), optimista: _presNum(vals[i][3]), basePct: _presNum(vals[i][4]),
                   seleccionado: String(vals[i][5] || '').trim().toLowerCase() || 'base' };
          break;
        }
      }
    }
    function pick(v, d) { return (v === undefined || v === null || v === '') ? d : v; }
    var cons = Math.max(0, _presNum(pick(body.conservador, prev.conservador)));
    var opt  = Math.max(0, _presNum(pick(body.optimista,   prev.optimista)));
    var base = _presNum(pick(body.basePct, prev.basePct));   // puede ser negativo (Base bajo el recomendado)
    var sel  = String(pick(body.seleccionado, prev.seleccionado) || 'base').trim().toLowerCase();
    if (['conservador', 'base', 'optimista', 'meta'].indexOf(sel) < 0) sel = 'base';
    var fila = [per, '__PROYECCION__', cons, opt, base, sel, ''];
    if (found) sh.getRange(found, 1, 1, fila.length).setValues([fila]);
    else sh.appendRow(fila);
    try { logAudit(body.usuario || '', 'Presupuesto', found ? 'Edición proyección' : 'Alta proyección',
      per + ' · esc=' + sel + ' base%=' + base + ' cons%=' + cons + ' opt%=' + opt, 'proyeccion', '', ''); } catch (e) {}
    return { ok: true, message: 'Proyección guardada.',
      proyeccion: { conservador: cons, optimista: opt, basePct: base, seleccionado: sel } };
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
      } else if (lin.toUpperCase() === '__PROYECCION__') {
        // Fila especial: parámetros de PROYECCIÓN (escenarios) por periodo — NO es una línea de meta,
        // no suma al total. col C = Conservador % (vs base), col D = Optimista % (vs base),
        // col E = Base % (vs recomendado del modelo), col F (Notas) = escenario seleccionado.
        map[per]._proyeccion = {
          conservador: _presNum(raw[r][2]),
          optimista: _presNum(raw[r][3]),
          basePct: _presNum(raw[r][4]),
          seleccionado: String(raw[r][5] || '').trim().toLowerCase() || 'base'
        };
      } else if (lin.substring(0, 3).toUpperCase() === 'EG:') {
        // Meta de EGRESO (prefijo EG:)
        map[per]._egLineas[lin.substring(3)] = meta;
        map[per].__totalEg += meta;
      } else {
        map[per]._lineas[lin] = { metaIngreso: meta, metaMargen: _presNum(raw[r][3]), crecObjetivo: _presNum(raw[r][4]) / 100, metaCantidad: _presNum(raw[r][6]) };
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
/* Etiqueta CANÓNICA de un grupo de ingresos — MISMA lógica que summary.gs
   (_summaryEsAgencia / _summaryEsExterno / _sumNorm), para que Presupuesto, Summary
   y Board hablen la misma taxonomía.
   Antes la columna U viajaba CRUDA (solo .trim()): "Externos" y "EXTERNOS" eran dos
   grupos distintos que la tabla pintaba idénticos (grupo duplicado en pantalla).
   `reg` es el registro clave-normalizada → etiqueta de display: la primera variante
   vista gana, así el nombre mostrado sigue siendo legible (no se aplasta a minúsculas). */
function _presCanonGrupo(l1, categoria, reg) {
  try {
    if (typeof _summaryEsAgencia === 'function' && (_summaryEsAgencia(l1) || _summaryEsAgencia(categoria))) return 'Agencias';
    if (typeof _summaryEsExterno === 'function' && (_summaryEsExterno(l1) || _summaryEsExterno(categoria))) return 'Externos';
  } catch (e) {}
  var k = (typeof _sumNorm === 'function') ? _sumNorm(l1) : String(l1 || '').trim().toLowerCase();
  if (!reg) return l1;
  if (!reg[k]) reg[k] = l1;
  return reg[k];
}
function _presIngresosProy(tY, tQ) {
  var q = {}, qCantTot = {}, origenSub = {};   // origenSub[gs] = true si el sub proviene de un ORIGEN externo atribuido
  var gReg = {};   // registro norm→display: deduplica la columna U ("Externos" == "EXTERNOS")
  var thisYear = new Date().getFullYear();
  for (var y = thisYear - 2; y <= thisYear; y++) {
    var rows; try { rows = _summaryReadIngresos(y); } catch (e) { rows = []; }
    rows.forEach(function (r) {
      var f = (r.fecha || '').substring(0, 10); if (f.length < 7) return;
      var yy = parseInt(f.substring(0, 4), 10), mo = parseInt(f.substring(5, 7), 10);
      var qk = _presQKey(yy, _presQ(mo));
      // Taxonomía CANÓNICA, alineada con summary.gs (_summaryEsAgencia/_summaryEsExterno):
      // agencias → "Agencias" (sub = nombre de la agencia u origen); externos → "Externos";
      // el resto conserva su nombre pero se deduplica sin distinguir mayúsculas.
      var _l1 = String(r.grupoU || '').trim() || String(r.categoria || '').trim() || '(Sin grupo)';
      var _l2 = String(r.categoria || '').trim();
      var g = _presCanonGrupo(_l1, r.categoria, gReg);
      if (g === 'Agencias') {
        var _agN = ''; try { _agN = _summaryAgenciaNombre(_l1) || _summaryAgenciaNombre(r.categoria); } catch (e) {}
        _l2 = _agN || _l2 || _l1;   // Agencias › REPROVIDA (nombre canónico) › items
      }
      // Ingresos externos atribuidos: el sub-nivel se abre por DUEÑO (Origen
      // externo). Los no-externos traen r.origen='' → sub = categoría (sin cambio).
      var _org = String(r.origen || '').trim();
      var _tieneOrigen = !!_org;
      var s = (g === 'Agencias') ? (_org || _l2 || g)
            : (g === 'Externos') ? (_org || 'Externos — sin atribuir')
            : (_l2 || g);
      if (_tieneOrigen) origenSub[g + '' + s] = true;   // este sub SÍ tiene dueño (origen) atribuido
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
  // ¿El grupo es el bucket de EXTERNOS? (ciclo/categoría con 'extern'). Sus sub-líneas se
  // muestran por DUEÑO: con origen → nombre del origen; sin origen → "Externos — sin atribuir".
  function _esExternos(name) { return /extern/i.test(String(name || '')); }
  Object.keys(allG).forEach(function (g) {
    var esExtGrupo = _esExternos(g);
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
      // Etiqueta a nivel medio de detalle (más que la categoría, menos que Summary):
      // externo con dueño → nombre del origen (= s); externo sin dueño → "sin atribuir".
      var _tuvoOrigen = !!origenSub[g + '' + s];
      var _extSin = esExtGrupo && !_tuvoOrigen;
      var _subLabel = _extSin ? 'Externos — sin atribuir' : s;
      subgrupos.push({ sub: s, subLabel: _subLabel, externoSinAtribuir: _extSin, categoriaOriginal: s,
        importeProy: sImp, cantProy: sCant, productos: productos });
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
    var esHistorico = false, _pedido = false;
    if (periodo && /^\d{4}-Q[1-4]$/.test(String(periodo).trim())) {
      var pm = String(periodo).trim().match(/^(\d{4})-Q([1-4])$/);
      tgtY = parseInt(pm[1], 10); tgtQ = parseInt(pm[2], 10); _pedido = true;
      // histórico = SOLO si el trimestre YA TERMINÓ. El trimestre EN CURSO y los
      // futuros son editables (permite ajustar el Q actual apenas empezado).
      // Un periodo ABIERTO manualmente (candado) también se vuelve editable.
      esHistorico = ((tgtY < curY) || (tgtY === curY && tgtQ < curQ)) && !_presPeriodoAbierto(String(periodo).trim());
    }
    // El tablero de META sigue al trimestre CONSULTADO. Antes quedaba clavado al
    // trimestre en curso por reloj e ignoraba `periodo` → se veía Q3 arriba y Q4 abajo.
    // Sin `periodo` se conserva el comportamiento por defecto (Q en curso arriba).
    var actY = _pedido ? tgtY : curY, actQ = _pedido ? tgtQ : curQ;
    var perActual = _presQKey(actY, actQ);
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

    // ── Trimestre del tablero (el CONSULTADO): pace ──
    var qStart = new Date(actY, (actQ - 1) * 3, 1);
    var qEnd = new Date(actY, actQ * 3, 0);
    var diasTotales = Math.round((qEnd - qStart) / 86400000) + 1;
    var _hoy0 = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    var _esFuturo = _hoy0 < qStart;   // trimestre por comenzar: no hay avance que medir
    var _esPasado = _hoy0 > qEnd;     // trimestre cerrado: transcurrió completo
    var diasTransc = _esFuturo ? 0 : (_esPasado ? diasTotales
      : Math.min(Math.max(Math.round((_hoy0 - qStart) / 86400000) + 1, 1), diasTotales));
    var realActual = (histQ[perActual] && histQ[perActual].__total) || 0;
    var fraccion = diasTotales > 0 ? diasTransc / diasTotales : 0;
    // Sin días transcurridos NO hay proyección de cierre (extrapolar de 0 no significa nada):
    // el semáforo "vas / no vas", la proyección de cierre y la brecha no aplican todavía.
    var proyCierre = _esFuturo ? null : (fraccion > 0 ? realActual / fraccion : realActual);
    var metaActualTotal = (metas[perActual] && metas[perActual].__total) || 0;
    var cumplimiento = (_esFuturo || proyCierre === null) ? null
      : (metaActualTotal > 0 ? proyCierre / metaActualTotal : null);
    var semaforo = _esFuturo ? 'futuro'
      : (cumplimiento === null ? 'sin-meta' : (cumplimiento >= 1 ? 'verde' : (cumplimiento >= 0.9 ? 'amarillo' : 'rojo')));

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

    // ── Parámetros de PROYECCIÓN (escenarios) guardados por periodo ──
    // La recomendación del modelo (ingProd.totalImporte) es solo la SEMILLA. Si el usuario
    // guardó su Base % / Conservador % / Optimista % y el escenario activo, esos MANDAN.
    // Va ANTES del margen a propósito: el margen DEBE derivar del escenario activo.
    var _proyCfg = (metas[perSig] && metas[perSig]._proyeccion) || null;
    var _escSeed = _presEscenariosCfg();
    var _consP  = (_proyCfg && _proyCfg.conservador != null) ? _proyCfg.conservador : _escSeed.conservador;
    var _optP   = (_proyCfg && _proyCfg.optimista  != null) ? _proyCfg.optimista  : _escSeed.optimista;
    var _basePct = (_proyCfg && _proyCfg.basePct   != null) ? _proyCfg.basePct    : 0;
    var _recomBase = ingProd.totalImporte || totProy;                              // recomendación cruda del modelo
    var _baseAjust = Math.round(_recomBase * (1 + (_basePct || 0) / 100));         // Base con el % vs recom del usuario
    var _consAmt   = Math.round(_baseAjust * (1 - (_consP || 0) / 100));
    var _optAmt    = Math.round(_baseAjust * (1 + (_optP  || 0) / 100));
    var _metaTot   = (metas[perSig] && metas[perSig].__total) || 0;
    // Escenario activo: el guardado manda; si no hay pero sí hay meta capturada, se implica 'meta'; si no, 'base'.
    var _selScen = (_proyCfg && _proyCfg.seleccionado) ? _proyCfg.seleccionado : (_metaTot > 0 ? 'meta' : 'base');
    var _proySel = _selScen === 'conservador' ? _consAmt
                 : _selScen === 'optimista'   ? _optAmt
                 : _selScen === 'meta'        ? _metaTot
                 : _baseAjust;

    // ════ FUENTE ÚNICA DE VERDAD de la pantalla ════
    //   INGRESO_DEL_PERIODO = proyeccionSeleccionada (el escenario ACTIVO)
    //   EGRESO_DEL_PERIODO  = egresos.proyeccion
    //   UTILIDAD            = INGRESO − EGRESO
    //   MARGEN %            = UTILIDAD / INGRESO
    // Antes el margen usaba la recomendación CRUDA del modelo (ingProd.totalImporte) y el
    // escenario se resolvía 35 líneas DESPUÉS → la Ganancia ignoraba el % de la Base y el
    // escenario elegido (se mostraba la utilidad del recomendado, no la del escenario activo).
    var _ingProyTot = _proySel;
    var margenProy = _ingProyTot - egProy;
    var margenPct = _ingProyTot > 0 ? (margenProy / _ingProyTot) * 100 : 0;

    // Adjuntar budget guardado (meta por producto) a la estructura nueva de ingresos.
    // La meta MANUAL manda: se devuelve el importe guardado (p.budget), la CANTIDAD
    // capturada (p.budgetCant) y una bandera de presencia (p.tieneMeta). El estimado
    // del sistema es solo la SEMILLA cuando NO hay meta guardada. Índice normalizado
    // (trim + minúsculas) por si el nombre del producto llegó con diferencias de
    // espacios/mayúsculas entre la captura y la relectura (evita reversión al estimado).
    var _metaNorm = {};
    Object.keys(metaSig).forEach(function (k) { _metaNorm[String(k).trim().toLowerCase()] = metaSig[k]; });
    function _metaDeProducto(prod) {
      var m = metaSig[prod]; if (m) return m;
      return _metaNorm[String(prod || '').trim().toLowerCase()] || null;
    }
    (ingProd.grupos || []).forEach(function (G) { (G.subgrupos || []).forEach(function (S) { (S.productos || []).forEach(function (p) {
      var mm = _metaDeProducto(p.producto);
      p.budget = (mm && mm.metaIngreso) || 0;
      p.budgetCant = (mm && mm.metaCantidad) || 0;
      p.tieneMeta = !!mm;                       // hay meta manual guardada (aunque valga 0)
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
        // esFuturo: el Q aún no arranca → sin semáforo, sin proyección de cierre, sin brecha.
        esFuturo: _esFuturo, esPasado: _esPasado,
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
        totales: { anioAnterior: totAnioAnt, reciente: totReciente, base: totBase,
            recomendado: _recomBase,                 // recomendación cruda del modelo (semilla)
            conservador: _consAmt,
            proyeccion: _baseAjust,                  // Base (con el % vs recom guardado por el usuario)
            optimista: _optAmt,
            meta: _metaTot,
            seleccionado: _selScen,                  // escenario activo (conservador|base|optimista|meta)
            proyeccionSeleccionada: _proySel }       // importe del escenario activo → lo consume el Board Deck
      },
      // %s + escenario activo del panel; esGuardado=true una vez que el usuario los persistió (no revertir al recom).
      escenariosPct: { conservador: _consP, optimista: _optP, basePct: _basePct, seleccionado: _selScen, esGuardado: !!_proyCfg },
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

/* ════════════════════════════════════════════════════════════════════════
   CONEXIÓN Presupuesto → Operating P&L
   Inyecta el Budget del módulo Presupuesto en el reporte P&L (readOperatingPL),
   reemplazando la pestaña "Budget" manual. Se llama desde core.gs después de
   readOperatingPL. Es DEFENSIVO: ante cualquier problema deja el P&L igual, y
   solo sobrescribe la línea cuando tiene un valor confiable del Presupuesto
   (las líneas sin match conservan su valor previo, no se borran).

   El P&L muestra un MES (targetMonths=[m]) o un TRIMESTRE completo (Q1..Q4),
   siempre dentro de un solo trimestre → budget = presupuesto del Q × (meses/3).
   ════════════════════════════════════════════════════════════════════════ */
function _presInyectaBudgetEnPL(plData, viewType, plMonth, plYear) {
  try {
    if (!plData || !plData.rows || !plData.rows.length) return plData;
    if (typeof readPresupuesto !== 'function') return plData;

    var MES_NUM = { Enero:1,Febrero:2,Marzo:3,Abril:4,Mayo:5,Junio:6,
                    Julio:7,Agosto:8,Septiembre:9,Octubre:10,Noviembre:11,Diciembre:12 };
    var QTR = { Q1:[1,2,3], Q2:[4,5,6], Q3:[7,8,9], Q4:[10,11,12] };
    viewType = String(viewType || 'Q1').trim();
    plMonth  = String(plMonth  || '').trim();
    var yr = parseInt(plYear || new Date().getFullYear(), 10) || new Date().getFullYear();
    var targetMonths = (plMonth && MES_NUM[plMonth]) ? [MES_NUM[plMonth]] : (QTR[viewType] || QTR.Q1);
    var qNum = _presQ(targetMonths[0]);            // 1..4
    var factor = targetMonths.length / 3;          // mes = 1/3 del trimestre
    var per = yr + '-Q' + qNum;

    var pp = readPresupuesto(per);
    if (!pp || !pp.ok) return plData;

    function norm(s){ return String(s||'').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/&/g,'and').replace(/\s+/g,' '); }

    // Budget trimestral por clave normalizada (meta si está fijada, si no la proyección).
    // OJO: subtipo (bud) y subgrupo (budSub) en mapas SEPARADOS para no duplicar cuando
    // el subtipo y su subgrupo se llaman igual (ej. "Medicamentos").
    var bud = {}, budG = {}, budSub = {}, revTot = 0, secTot = { COGS:0, OPEX:0, GA:0, TAXES:0 };
    // INGRESOS POR GRUPO (columna U: Alta/Surrogacy/Externos/Other Income) — es la dimensión
    // que usa el P&L. Antes comparaba contra categorías (nivel fino) y por eso Alta/Surrogacy
    // no calzaban (su categoría interna tiene otro nombre). revTot = total de la proyección.
    var _grps = (pp.siguiente && pp.siguiente.ingresosGrupos) || [];
    _grps.forEach(function(g){ if (g && g.grupo) budG[norm(g.grupo)] = (budG[norm(g.grupo)] || 0) + (Number(g.importeProy)||0); });
    revTot = _grps.reduce(function(s,g){ return s + (Number(g.importeProy)||0); }, 0);
    // También por CATEGORÍA (nivel fino) por si alguna línea calza mejor ahí o hay meta por línea
    (pp.siguiente && pp.siguiente.lineas || []).forEach(function(l){
      var v = (Number(l.meta) > 0 ? Number(l.meta) : Number(l.proyeccion)) || 0;
      if (l.linea) bud[norm(l.linea)] = (bud[norm(l.linea)] || 0) + v;
    });
    if (revTot <= 0) { (pp.siguiente && pp.siguiente.lineas || []).forEach(function(l){ revTot += (Number(l.meta)>0?Number(l.meta):Number(l.proyeccion))||0; }); }
    (pp.egresos && pp.egresos.lineas || []).forEach(function(l){
      var v = (Number(l.meta) > 0 ? Number(l.meta) : Number(l.proyeccion)) || 0;
      if (l.linea)    bud[norm(l.linea)]       = (bud[norm(l.linea)]       || 0) + v;   // subtipo
      if (l.subgrupo) budSub[norm(l.subgrupo)] = (budSub[norm(l.subgrupo)] || 0) + v;   // subgrupo (Payroll…)
      var sec = String(l.seccion || '').toUpperCase();
      if (secTot[sec] != null) secTot[sec] += v;
    });
    if (revTot <= 0 && secTot.COGS <= 0 && secTot.OPEX <= 0 && secTot.GA <= 0) return plData; // sin datos → no tocar

    // Métricas derivadas (base devengado del P&L)
    var mGross = revTot - secTot.COGS;
    var mClinic = mGross - secTot.OPEX;
    var mEbitda = mClinic - secTot.GA;
    var mNet = mEbitda - secTot.TAXES;

    // Alias P&L (español) → subgrupo del Presupuesto (inglés). Si el subgrupo no existe,
    // simplemente no matchea y la línea conserva su valor previo (inofensivo).
    var ALIAS = { 'nomina':'payroll', 'renta':'rent and facilities', 'mto renta':'rent and facilities',
      'servicios':'software and services', 'marketing':'marketing and advertising', 'gastos varios':'other ganda' };

    // Resolver el budget de una fila del P&L por su etiqueta (devuelve número o null)
    function resolveLinea(label){
      var k = norm(label);
      switch (k) {
        case 'revenue': case 'total income': return revTot;
        case 'cogs': case 'total cogs': return secTot.COGS;
        case 'opex': case 'total opex': return secTot.OPEX;
        case 'ganda': case 'total ganda': case 'ga': return secTot.GA;   // "g&a" → "ganda"
        case 'taxes': return secTot.TAXES;
        case 'gross profit': return mGross;
        case 'clinic contribution': return mClinic;
        case 'ebitda': return mEbitda;
        case 'net profit': return mNet;
      }
      // Agencias: suma el budget de todos los grupos de ingreso que sean agencia.
      if (k === 'agencias' && typeof _summaryEsAgencia === 'function') {
        var _ta = 0; for (var _gk in budG) { if (_summaryEsAgencia(_gk)) _ta += budG[_gk]; } return _ta || null;
      }
      if (budG[k] != null) return budG[k];               // GRUPO de ingreso (Alta/Surrogacy/Externos/Other Income)
      if (bud[k] != null) return bud[k];                 // subtipo / categoría exacta
      if (budSub[k] != null) return budSub[k];           // subgrupo exacto
      if (ALIAS[k]) { if (bud[ALIAS[k]] != null) return bud[ALIAS[k]]; if (budSub[ALIAS[k]] != null) return budSub[ALIAS[k]]; }
      // ingresos: match laxo por substring, primero sobre GRUPOS, luego categorías
      var gKeys = Object.keys(budG);
      for (var i=0;i<gKeys.length;i++){ var gk=gKeys[i]; if(gk && (gk===k || gk.indexOf(k)>=0 || k.indexOf(gk)>=0)) return budG[gk]; }
      var cKeys = (pp.siguiente && pp.siguiente.lineas || []).map(function(l){ return norm(l.linea); });
      for (var j=0;j<cKeys.length;j++){ var ck=cKeys[j]; if(ck && (ck===k || ck.indexOf(k)>=0 || k.indexOf(ck)>=0)) return bud[ck]; }
      return null;   // sin match → conservar valor previo
    }

    var totBudget = revTot * factor;   // denominador para %Budget
    plData.rows.forEach(function(row){
      if (!row.values || row.values.length < 5) return;
      var b = resolveLinea(row.label);
      // Tratar 0/null como "sin dato del Presupuesto" → NO pisar el valor previo con 0
      if (b === null || b === undefined || b === 0) {
        var bgOld = Number(row.values[2]) || 0;
        row.values[3] = totBudget ? bgOld / totBudget : null;
        row.values[4] = bgOld ? (Number(row.values[0]) - bgOld) / Math.abs(bgOld) : null;
        return;
      }
      var bg = b * factor;
      row.values[2] = bg;
      row.values[3] = totBudget ? bg / totBudget : null;
      row.values[4] = bg ? (Number(row.values[0]) - bg) / Math.abs(bg) : null;
    });

    // Marcar la fuente en el encabezado y en un flag para el frontend
    if (plData.colHeaders && plData.colHeaders[2]) plData.colHeaders[2].label += ' · Presupuesto';
    // Diagnóstico: GRUPOS de ingreso del Presupuesto (dimensión del P&L) para revisar el mapeo
    var cats = _grps.map(function(g){
      return { n:String(g.grupo||''), m: Math.round((Number(g.importeProy)||0) * factor) };
    }).filter(function(x){ return x.m>0; }).sort(function(a,b){ return b.m-a.m; });
    plData.budgetFuente = { origen:'presupuesto', periodo:per, factor:factor,
      mensual:(targetMonths.length===1), nota:(targetMonths.length===1?'Budget mensual = presupuesto trimestral ÷ 3':'Budget del trimestre'),
      categorias: cats };
    return plData;
  } catch (ex) { return plData; }   // nunca romper el P&L
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
    var fila = [per, lin, _presNum(body.metaIngreso), _presNum(body.metaMargen), _presNum(body.crecimiento), String(body.notas || ''), _presNum(body.metaCantidad)];
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
      var row = [per, lin, _presNum(m.metaIngreso), 0, 0, '', _presNum(m.metaCantidad)];
      var found = existing[lin.toLowerCase()];
      if (found) sh.getRange(found, 1, 1, row.length).setValues([row]);
      else appends.push(row);
    });
    if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, appends[0].length).setValues(appends);
    return { ok: true, message: body.metas.length + ' metas guardadas.' };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ════════════════════════════════════════════════════════════════════════
   ETAPA 1 — MOTOR DE COSTOS FIJO/VARIABLE + P&L (utilidad pronosticada)
   Hoja Presupuesto_Costos: Línea | Sección | Tipo | Modo | Valor | Notas
     Sección ∈ COGS/OPEX/GA/TAXES  (dónde cae el gasto en el P&L)
     Tipo    ∈ FIJO | VARIABLE
     Modo    ∈ monto (mensual, para FIJO) | pctIngreso (% de ingreso) | porTratamiento ($/tratamiento)
   Al mover ingresos/tratamientos, los VARIABLES se recalculan → utilidad real.
   ════════════════════════════════════════════════════════════════════════ */
var PRES_COSTOS_TAB = 'Presupuesto_Costos';
var PRES_COSTOS_HEADERS = ['Línea (subtipo)', 'Sección', 'Tipo', 'Modo', 'Valor', 'Notas'];
function _presCostosSeed(sh) {
  // Semilla desde los subtipos reales de egresos: FIJO monto = promedio MENSUAL del último trimestre.
  try {
    var eg = _presHistoricoEgresos();
    var hoy = new Date(), cy = hoy.getFullYear(), cq = _presQ(hoy.getMonth() + 1);
    var lc = _presPrevQ(cy, cq), kRec = _presQKey(lc.y, lc.q);
    var rows = [];
    (eg.subtipos || []).forEach(function (sub) {
      var qTot = (eg.q[kRec] && eg.q[kRec][sub]) || 0;
      var mensual = Math.round(qTot / 3);
      var cont = eg.contableBySub[sub] || 'Gasto';
      var seccion = 'OPEX';
      try {
        var c = _summaryDefaultClass('egreso', cont + '|' + sub);
        var subg = _summaryEgSubgroup(sub, c.grupo);
        seccion = SUMMARY_EG_SUBGROUP_SECTION[_sumNorm(subg)] || c.grupo || 'OPEX';
      } catch (e) {}
      rows.push([sub, seccion, 'FIJO', 'monto', mensual, '']);
    });
    if (rows.length) sh.getRange(2, 1, rows.length, PRES_COSTOS_HEADERS.length).setValues(rows);
  } catch (e) {}
}
function _presEnsureCostos() {
  var ss = SpreadsheetApp.openById(ER_SS_ID);
  var sh = ss.getSheetByName(PRES_COSTOS_TAB);
  if (!sh) {
    sh = ss.insertSheet(PRES_COSTOS_TAB);
    sh.getRange(1, 1, 1, PRES_COSTOS_HEADERS.length).setValues([PRES_COSTOS_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
    sh.setFrozenRows(1);
    var w = [220, 90, 100, 130, 110, 200]; for (var i = 0; i < w.length; i++) sh.setColumnWidth(i + 1, w[i]);
    _presCostosSeed(sh);
  }
  return sh;
}
function readPresupuestoCostos() {
  try {
    var sh = _presEnsureCostos();
    var raw = sh.getDataRange().getValues(); var out = [];
    for (var i = 1; i < raw.length; i++) {
      var r = raw[i]; var lin = String(r[0] || '').trim(); if (!lin) continue;
      out.push({ linea: lin, seccion: String(r[1] || 'OPEX').trim().toUpperCase(), tipo: String(r[2] || 'FIJO').trim().toUpperCase(),
        modo: String(r[3] || 'monto').trim(), valor: _presNum(r[4]), notas: String(r[5] || '') });
    }
    return { ok: true, costos: out, _setup: true };
  } catch (ex) { return { ok: false, error: ex.message, costos: [] }; }
}
function savePresupuestoCostos(body) {
  try {
    if (!body || !body.costos || !body.costos.length) return { ok: false, error: 'Datos incompletos.' };
    var sh = _presEnsureCostos(); sh.clear();
    sh.getRange(1, 1, 1, PRES_COSTOS_HEADERS.length).setValues([PRES_COSTOS_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
    var rows = [];
    body.costos.forEach(function (c) {
      var lin = String(c.linea || '').trim(); if (!lin) return;
      rows.push([lin, String(c.seccion || 'OPEX').toUpperCase(), String(c.tipo || 'FIJO').toUpperCase(), String(c.modo || 'monto'), _presNum(c.valor), String(c.notas || '')]);
    });
    if (rows.length) sh.getRange(2, 1, rows.length, PRES_COSTOS_HEADERS.length).setValues(rows);
    try { logAudit(body.usuario || '', 'Presupuesto', 'Guardar costos', '', '', '', rows.length + ' líneas'); } catch (e) {}
    return { ok: true, guardados: rows.length };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
// Proyecta los egresos por sección desde la estructura fijo/variable.
// ingresoQ = ingreso del trimestre; tratQ = nº de tratamientos del trimestre.
function _presCostosProy(costos, ingresoQ, tratQ) {
  var sec = { COGS: 0, OPEX: 0, GA: 0, TAXES: 0 }, lineas = [];
  (costos || []).forEach(function (c) {
    var m = 0;
    if (c.tipo === 'VARIABLE') {
      if (c.modo === 'pctIngreso') m = ingresoQ * (_presNum(c.valor) / 100);
      else if (c.modo === 'porTratamiento') m = _presNum(c.valor) * (tratQ || 0);
      else m = _presNum(c.valor) * 3;
    } else { m = _presNum(c.valor) * 3; }   // FIJO monto mensual → trimestre = ×3
    var s = (c.seccion || 'OPEX').toUpperCase(); if (sec[s] == null) s = 'OPEX';
    sec[s] += m;
    lineas.push({ linea: c.linea, seccion: s, tipo: c.tipo, modo: c.modo, montoQ: m });
  });
  return { secciones: sec, total: sec.COGS + sec.OPEX + sec.GA + sec.TAXES, lineas: lineas };
}
// P&L en vivo desde la estructura de costos (motor etapa 1).
function readPresupuestoModelo(periodo) {
  try {
    var hoy = new Date(), cy = hoy.getFullYear(), cq = _presQ(hoy.getMonth() + 1);
    var nx = _presNextQ(cy, cq); var tY = nx.y, tQ = nx.q;
    if (periodo && /^\d{4}-Q[1-4]$/.test(String(periodo).trim())) {
      var pm = String(periodo).trim().match(/^(\d{4})-Q([1-4])$/); tY = parseInt(pm[1], 10); tQ = parseInt(pm[2], 10);
    }
    var ing = _presIngresosProy(tY, tQ);
    var ingresoQ = ing.totalImporte || 0;
    var tratQ = 0; (ing.grupos || []).forEach(function (g) { tratQ += g.cantProy || 0; });
    var costos = (readPresupuestoCostos().costos) || [];
    var proy = _presCostosProy(costos, ingresoQ, tratQ);
    var s = proy.secciones;
    var utilBruta = ingresoQ - s.COGS;
    var utilClinica = utilBruta - s.OPEX;
    var ebitda = utilClinica - s.GA;
    var utilNeta = ebitda - s.TAXES;
    function pct(n) { return ingresoQ > 0 ? n / ingresoQ * 100 : 0; }
    return {
      ok: true, periodo: _presQKey(tY, tQ), ingresos: ingresoQ, tratamientos: tratQ,
      cogs: s.COGS, opex: s.OPEX, ga: s.GA, taxes: s.TAXES, costoTotal: proy.total,
      utilidadBruta: utilBruta, utilidadClinica: utilClinica, ebitda: ebitda, utilidadNeta: utilNeta,
      margenBruto: pct(utilBruta), margenClinica: pct(utilClinica), margenEbitda: pct(ebitda), margenNeto: pct(utilNeta),
      lineas: proy.lineas
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
