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

/* ── Lee histórico de egresos por trimestre (total) ─────────────── */
function _presHistoricoEgresos() {
  var q = {};
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
      var iF = col('fecha'), iE = col('egresos'), iV = col('vencimiento');
      for (var r = 1; r < raw.length; r++) {
        var row = raw[r];
        var f = (iF > -1 ? row[iF] : '') || (iV > -1 ? row[iV] : '');
        var d = (f instanceof Date) ? f : new Date(f);
        if (!d || isNaN(d.getTime())) continue;
        var y = d.getFullYear(), qq = _presQ(d.getMonth() + 1);
        var monto = _presNum(iE > -1 ? row[iE] : 0);
        if (!monto) continue;
        var qk = _presQKey(y, qq);
        q[qk] = (q[qk] || 0) + monto;
      }
    } catch (e) {}
  });
  return q;
}

/* ── Lee metas de la hoja ───────────────────────────────────────── */
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
      if (!map[per]) map[per] = { __total: 0, _lineas: {} };
      var meta = _presNum(raw[r][2]);
      map[per]._lineas[lin] = {
        metaIngreso: meta,
        metaMargen: _presNum(raw[r][3]),
        crecObjetivo: _presNum(raw[r][4]) / 100
      };
      map[per].__total += meta;
    }
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
function readPresupuesto() {
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
    var perActual = _presQKey(curY, curQ);
    var perSig = _presQKey(tgtY, tgtQ);

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

    // ── Tendencia mensual: últimos 12 meses reales + 3 proyectados ──
    var tendencia = _presTendencia(histM, curY, curQ, tgtY, tgtQ, totProy, histQ, kAnioAnt);

    // ── Egresos + margen (nivel total, mismo ratchet) ──
    var egAnioAnt = egQ[kAnioAnt] || 0, egReciente = egQ[kReciente] || 0;
    var gEg = _presCrecimientoEgresos(egQ, curY, curQ);
    var egBase = egAnioAnt > 0 ? egAnioAnt : egReciente;
    var gEgAplicado = Math.min(Math.max(gEg, 0), PRES_CREC_MAX);
    var egProy = Math.max(egBase * (1 + gEgAplicado), Math.max(egAnioAnt, egReciente));
    var margenProy = totProy - egProy;
    var margenPct = totProy > 0 ? (margenProy / totProy) * 100 : 0;

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
      siguiente: {
        periodo: perSig,
        lineas: lineasProy,
        totales: { anioAnterior: totAnioAnt, reciente: totReciente, base: totBase, conservador: totCons, proyeccion: totProy, optimista: totOpt, meta: totMeta }
      },
      egresos: { anioAnterior: egAnioAnt, reciente: egReciente, proyeccion: egProy, margenProyectado: margenProy, margenPct: margenPct },
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

function _presCrecimientoEgresos(egQ, refY, refQ) {
  var ult = 0, prev = 0, yy = refY, qq = refQ;
  for (var i = 0; i < 4; i++) { var p = _presPrevQ(yy, qq); ult += egQ[_presQKey(p.y, p.q)] || 0; yy = p.y; qq = p.q; }
  for (var j = 0; j < 4; j++) { var p2 = _presPrevQ(yy, qq); prev += egQ[_presQKey(p2.y, p2.q)] || 0; yy = p2.y; qq = p2.q; }
  if (prev <= 0) return 0;
  return (ult - prev) / prev;
}

/* ── Serie mensual para la gráfica ──────────────────────────────── */
function _presTendencia(histM, curY, curQ, tgtY, tgtQ, totProySig, histQ, kAnioAnt) {
  var MESES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  // Construir 12 meses hacia atrás desde el mes actual + 3 meses del trimestre objetivo
  var meses = [], real = [], proy = [];
  var hoy = new Date();
  var y = hoy.getFullYear(), mo = hoy.getMonth() + 1;
  // arrancar 9 meses atrás
  var startY = y, startM = mo - 9;
  while (startM <= 0) { startM += 12; startY--; }
  var yy = startY, mm = startM;
  for (var i = 0; i < 9; i++) {
    var mk = yy + '-' + String(mm).padStart(2, '0');
    meses.push(MESES[mm - 1] + ' ' + String(yy).slice(2));
    real.push((histM[mk] && histM[mk].__total) || 0);
    proy.push(null);
    mm++; if (mm > 12) { mm = 1; yy++; }
  }
  // Distribuir la proyección del trimestre objetivo entre sus 3 meses,
  // usando la forma mensual del mismo trimestre del año anterior.
  var qMonths = [(tgtQ - 1) * 3 + 1, (tgtQ - 1) * 3 + 2, (tgtQ - 1) * 3 + 3];
  var shapePrev = qMonths.map(function (m2) { var mk2 = (tgtY - 1) + '-' + String(m2).padStart(2, '0'); return (histM[mk2] && histM[mk2].__total) || 0; });
  var shapeSum = shapePrev.reduce(function (a, b) { return a + b; }, 0);
  var lastReal = real[real.length - 1] || 0;
  // conectar la línea de proyección desde el último real
  proy[proy.length - 1] = lastReal;
  qMonths.forEach(function (m2, idx) {
    meses.push(MESES[m2 - 1] + ' ' + String(tgtY).slice(2));
    real.push(null);
    var parte = shapeSum > 0 ? (shapePrev[idx] / shapeSum) : (1 / 3);
    proy.push(totProySig * parte);
  });
  return { meses: meses, real: real, proyeccion: proy };
}

/* ── Guardar / actualizar una meta ──────────────────────────────── */
function savePresupuestoMeta(body) {
  try {
    if (!body || !String(body.periodo || '').trim() || !String(body.linea || '').trim())
      return { ok: false, error: 'Periodo y línea son obligatorios.' };
    var ss = SpreadsheetApp.openById(ER_SS_ID);
    var sh = ss.getSheetByName(PRES_METAS_TAB);
    if (!sh) { setupPresupuesto(); sh = ss.getSheetByName(PRES_METAS_TAB); }
    var per = String(body.periodo).trim(), lin = String(body.linea).trim();
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
