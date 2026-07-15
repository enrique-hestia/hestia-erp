/* ==============================================================
   analisis.gs — Centro de Análisis de Inteligencia de Negocio
   --------------------------------------------------------------
   Fuente autoritativa para ingresos: Estado de Resultados (ER).
   Fuente operativa para mix/producto: BD_Ingresos.
   Fuente de plan: Budget tab en ER_SS_ID.

   Rutas GET:
     ?action=analisisIngresos   → readAnalisisIngresos()
     ?action=analisisEgresos    → readAnalisisEgresos()
     ?action=analisisPacientes  → readAnalisisPacientes()
     ?action=analisisServicios  → readAnalisisServicios()
     ?action=analisisSurrogacy  → readAnalisisSurrogacy()
     ?action=analisisRentabilidad → readAnalisisRentabilidad()
   ============================================================== */

function _anNum(v) { if (typeof v === 'number') return v; var n = parseFloat(String(v || '').replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
function _anFM(v) { v = Math.abs(Number(v) || 0); if (v >= 1e6) return '$' + (v / 1e6).toFixed(2) + 'M'; if (v >= 1e3) return '$' + Math.round(v / 1e3) + 'K'; return '$' + Math.round(v); }

/* ── Helpers comunes ──────────────────────────────────────────── */

// Lee "Total Income" y por-categoría del Estado de Resultados → {'2026-04': 1500000}
function _anReadErSheetMonthly(sh) {
  if (!sh) return { income: {}, byLine: {} };
  var MON = {jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'};
  var raw = sh.getDataRange().getValues();
  var hdr = raw.length > 1 ? raw[1] : [];
  var colMap = [];
  for (var j = 3; j < hdr.length; j++) {
    var v = hdr[j];
    if (v instanceof Date) {
      var yr = v.getFullYear(), mn = v.getMonth()+1;
      if (yr >= 2020 && yr <= 2035) colMap.push({idx:j, ym: yr+'-'+String(mn).padStart(2,'0')});
      continue;
    }
    var s = String(v||'').trim();
    var m2 = s.match(/^([A-Za-z]{3})-(\d{2})$/);   // Mmm-YY
    if (m2) { var mo2 = MON[m2[1].toLowerCase()]; if (mo2) { colMap.push({idx:j, ym:'20'+m2[2]+'-'+mo2}); continue; } }
    var m4 = s.match(/^([A-Za-z]{3})-(\d{4})$/);   // Mmm-YYYY
    if (m4) { var mo4 = MON[m4[1].toLowerCase()]; if (mo4) { colMap.push({idx:j, ym:m4[2]+'-'+mo4}); continue; } }
  }
  var income = {}, byLine = {}, inIncome = false;
  for (var ri = 4; ri < raw.length; ri++) {
    var r = raw[ri];
    var a = String(r[0]||'').trim(), b = String(r[1]||'').trim(), c = String(r[2]||'').trim();
    if (!a && !b && !c) continue;
    // Detect section start: col A alone = section header
    if (a && !b && !c) { inIncome = (a.toLowerCase() === 'income'); continue; }
    // Category row (col B)
    if (inIncome && b && !a && !c) {
      var catData = {};
      colMap.forEach(function(col) { var v2 = r[col.idx]; catData[col.ym] = (v2 instanceof Date) ? 0 : (Number(v2)||0); });
      byLine[b] = catData;
    }
    // Total Income row (col C, B, or A — buscar en las primeras 3 columnas)
    var _lbl = (c || b || a).toLowerCase().replace(/\s+/g,' ').trim();
    if (_lbl === 'total income' || _lbl === 'ingresos totales') {
      colMap.forEach(function(col) { var v2 = r[col.idx]; income[col.ym] = (v2 instanceof Date) ? 0 : (Number(v2)||0); });
      inIncome = false;
    }
  }
  return { income: income, byLine: byLine };
}

// Abre ER y Budget en una sola llamada (evita 2 open del mismo SS)
function _anLoadErBudget() {
  var ssEr = SpreadsheetApp.openById(ER_SS_ID);
  var erSh = null, bgSh = null, all = ssEr.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getSheetId() === ER_GID) erSh = all[i];
    if (all[i].getSheetId() === BUDGET_GID) bgSh = all[i];
  }
  if (!erSh) erSh = all[0];
  var er = _anReadErSheetMonthly(erSh);
  var bg = bgSh ? _anReadErSheetMonthly(bgSh) : { income: {}, byLine: {} };
  return { er: er, budget: bg.income, budgetByLine: bg.byLine };
}

// Lee BD_Ingresos y devuelve rollup por año/mes/cat/prod
function _anReadBDIngresos() {
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = null, all = ss.getSheets();
    for (var i = 0; i < all.length; i++) if (all[i].getName() === BD_INGRESOS_TAB) { sh = all[i]; break; }
    if (!sh) return null;
    var raw = sh.getDataRange().getValues();
    // cols: OP(0),Linea(1),Fecha(2),Paciente(3),Cat(4),Prod(5),PVP(6),Desc(7),Cant(8),TotalPagar(9),...
    var byCatYear = {}, byProd = {}, byPaciente = {}, byMonth = {}, opSet = {};
    for (var r = 1; r < raw.length; r++) {
      var row = raw[r];
      var op = String(row[0]||'').trim(); if (!op) continue;
      var f = row[2]; var d = (f instanceof Date) ? f : new Date(f);
      if (!d || isNaN(d.getTime())) continue;
      var y = d.getFullYear(), mo = d.getMonth()+1;
      var mk = y+'-'+String(mo).padStart(2,'0');
      var cat = String(row[4]||'').trim() || 'Sin categoría';
      var prod = String(row[5]||'').trim() || cat;
      var pac = String(row[3]||'').trim() || 'Sin nombre';
      var monto = _anNum(row[9]);
      var cant = _anNum(row[8]) || 0;
      var key = y+'|'+cat;
      if (!byCatYear[key]) byCatYear[key] = { cat:cat, y:y, total:0, cant:0, ops:{} };
      byCatYear[key].total += monto; byCatYear[key].cant += cant; byCatYear[key].ops[op] = 1;
      if (!byProd[prod]) byProd[prod] = { prod:prod, cat:cat, total:0, cant:0, ops:{} };
      byProd[prod].total += monto; byProd[prod].cant += cant; byProd[prod].ops[op] = 1;
      if (!byPaciente[pac]) byPaciente[pac] = { pac:pac, total:0, ops:{} };
      byPaciente[pac].total += monto; byPaciente[pac].ops[op] = 1;
      if (!byMonth[mk]) byMonth[mk] = {};
      byMonth[mk][cat] = (byMonth[mk][cat]||0) + monto;
      opSet[op] = { pac:pac, fecha:mk, cat:cat };
    }
    return { byCatYear:byCatYear, byProd:byProd, byPaciente:byPaciente, byMonth:byMonth, opSet:opSet };
  } catch(ex) { return null; }
}

/* ══════════════════════════════════════════════════════════════
   ANÁLISIS DE INGRESOS
   Fuente primaria: Estado de Resultados (ER) — números reales
   Fuente operativa: BD_Ingresos — mix por categoría/producto
   Fuente de plan: Budget tab
   ══════════════════════════════════════════════════════════════ */
function readAnalisisIngresos() {
  try {
    var anioActual = new Date().getFullYear();
    var anioAnterior = anioActual - 1;
    var hoy = new Date();
    var mesActual = anioActual+'-'+String(hoy.getMonth()+1).padStart(2,'0');

    // 1. ER + Budget (autoritativos)
    var eb = _anLoadErBudget();
    var erIncome = eb.er.income;   // {'2026-04': 2500000}
    var budget   = eb.budget;      // {'2026-04': 2800000}

    // 2. BD_Ingresos (mix operativo)
    var bd = _anReadBDIngresos();

    // 3. Serie mensual: combina ER real + Budget en el mismo array
    var allYms = {};
    Object.keys(erIncome).forEach(function(ym) { allYms[ym] = 1; });
    Object.keys(budget).forEach(function(ym)   { allYms[ym] = 1; });
    var mesesArr = Object.keys(allYms).sort().map(function(ym) {
      return { mes:ym, total: erIncome[ym]||0, budget: budget[ym]||0 };
    });

    // 4. YTD
    var ytdActual=0, ytdBudget=0, ytdPrev=0;
    Object.keys(erIncome).sort().forEach(function(ym) {
      var y = ym.substring(0,4), mStr = ym.substring(5,7);
      if (y == anioActual && ym <= mesActual) { ytdActual += erIncome[ym]||0; ytdBudget += budget[ym]||0; }
      if (y == anioAnterior && mStr <= mesActual.substring(5)) ytdPrev += erIncome[ym]||0;
    });
    // Budget completo del año (todos los meses de la pestaña Budget)
    var ytdBudgetFull = 0;
    Object.keys(budget).forEach(function(ym) { if (ym.substring(0,4) == anioActual) ytdBudgetFull += budget[ym]||0; });

    // 5. Mix por categoría desde BD_Ingresos (año actual)
    var curCat = {}, prevCat = {}, curTotal = 0, prevTotal = 0;
    if (bd) {
      Object.keys(bd.byCatYear).forEach(function(key) {
        var k = bd.byCatYear[key];
        if (k.y === anioActual) {
          curCat[k.cat] = k; curTotal += k.total;
        } else if (k.y === anioAnterior) {
          prevCat[k.cat] = (prevCat[k.cat]||0) + k.total; prevTotal += k.total;
        }
      });
    }
    var categorias = Object.keys(curCat).map(function(c) {
      var k=curCat[c], ant=prevCat[c]||0, ops=Object.keys(k.ops).length;
      return { categoria:c, total:k.total, pct:curTotal>0?k.total/curTotal*100:0, cantidad:k.cant, operaciones:ops, ticket:ops>0?k.total/ops:0, totalAnt:ant, crec:ant>0?((k.total-ant)/ant*100):null };
    }).sort(function(a,b){ return b.total-a.total; });

    // 6. Top productos
    var productos = bd ? Object.keys(bd.byProd).map(function(p) {
      var k=bd.byProd[p]; var ops=Object.keys(k.ops).length;
      return { producto:p, categoria:k.cat, total:k.total, cantidad:k.cant, ticket:ops>0?k.total/ops:0 };
    }).sort(function(a,b){ return b.total-a.total; }).slice(0,15) : [];
    var masVendidos = bd ? Object.keys(bd.byProd).map(function(p) {
      var k=bd.byProd[p]; return { producto:p, categoria:k.cat, total:k.total, cantidad:k.cant };
    }).filter(function(p){ return p.cantidad>0; }).sort(function(a,b){ return b.cantidad-a.cantidad; }).slice(0,12) : [];

    // 7. Pace + credibilidad
    var pctCumplimiento = ytdBudget>0 ? ytdActual/ytdBudget : null;
    var crecYoY = ytdPrev>0 ? (ytdActual-ytdPrev)/ytdPrev : null;
    var diasAno = anioActual%4===0?366:365;
    var diaActual = Math.floor((hoy - new Date(anioActual,0,0)) / 86400000);
    var paceAnualizado = diaActual>0 ? ytdActual*diasAno/diaActual : 0;
    var mesesRestantes = 11 - hoy.getMonth(); // meses completos restantes
    var budgetRestante = ytdBudgetFull - ytdActual;
    var recentArr = Object.keys(erIncome).filter(function(ym){ return ym.substring(0,4)==anioActual && ym<=mesActual; }).sort().slice(-3);
    var avgReciente = recentArr.length>0 ? recentArr.reduce(function(s,ym){ return s+(erIncome[ym]||0); },0)/recentArr.length : 0;
    var avgNecesario = mesesRestantes>0 ? budgetRestante/mesesRestantes : 0;
    var credibilidad = avgNecesario>0 && avgReciente>0 ? avgReciente/avgNecesario : null;
    var credTxt = credibilidad===null?'Sin datos suficientes':(credibilidad>=1.1?'Muy alcanzable':(credibilidad>=0.9?'Alcanzable':(credibilidad>=0.7?'Exigente':'Muy exigente')));
    var semaforo = pctCumplimiento===null?'sin-meta':(pctCumplimiento>=1?'verde':(pctCumplimiento>=0.9?'amarillo':'rojo'));

    // 8. Insights CEO
    var insights = [];
    if (pctCumplimiento!==null) {
      var gap = ytdBudget - ytdActual;
      insights.push({ tipo: pctCumplimiento>=1?'ok':(pctCumplimiento>=0.85?'warn':'warn'), icono: pctCumplimiento>=1?'check-circle':'alert-circle',
        titulo: 'Budget YTD: '+(pctCumplimiento*100).toFixed(0)+'% de cumplimiento',
        detalle: 'Real: '+_anFM(ytdActual)+' vs Plan: '+_anFM(ytdBudget)+(gap>0?' · Brecha: '+_anFM(gap)+' para cerrar el año':' · ¡Encima del plan!') });
    }
    if (crecYoY!==null) {
      insights.push({ tipo:crecYoY>=0?'ok':'warn', icono:crecYoY>=0?'trending-up':'trending-down',
        titulo: (crecYoY>=0?'+':'')+_anFM(ytdActual-ytdPrev)+' vs mismo período '+anioAnterior,
        detalle: 'Crecimiento real: '+(crecYoY>=0?'+':'')+( crecYoY*100).toFixed(1)+'% YoY · YTD '+anioAnterior+': '+_anFM(ytdPrev) });
    }
    if (credibilidad!==null) {
      insights.push({ tipo:credibilidad>=0.9?'ok':'warn', icono:'target',
        titulo: 'Budget restante: '+credTxt,
        detalle: 'Ritmo reciente '+_anFM(avgReciente)+'/mes vs '+_anFM(avgNecesario)+'/mes necesario · Faltan '+mesesRestantes+' meses' });
    }
    if (categorias.length) {
      insights.push({ tipo:'info', icono:'award', titulo:categorias[0].categoria+' lidera con '+categorias[0].pct.toFixed(0)+'% del mix',
        detalle: _anFM(categorias[0].total)+' en '+anioActual+' (datos operativos BD_Ingresos).' });
    }
    var numOps = bd ? Object.keys(bd.opSet).length : 0;

    return {
      ok:true, fuente:'Estado de Resultados (ER) + BD_Ingresos',
      anioActual:anioActual, anioAnterior:anioAnterior,
      ytdActual:ytdActual, ytdBudget:ytdBudget, ytdPrev:ytdPrev, ytdBudgetFull:ytdBudgetFull,
      crecYoY:crecYoY, pctCumplimiento:pctCumplimiento, semaforo:semaforo,
      totalActual:ytdActual,
      paceAnualizado:paceAnualizado, budgetRestante:budgetRestante,
      avgMensualReciente:avgReciente, avgMensualNecesario:avgNecesario,
      credibilidad:credibilidad, credibilidadTxt:credTxt,
      numOperaciones:numOps, ticketPromedio:numOps>0?curTotal/numOps:0,
      meses:mesesArr, categorias:categorias, topRevenue:productos, masVendidos:masVendidos, insights:insights
    };
  } catch(ex) { return { ok:false, error:ex.message }; }
}

/* ══════════════════════════════════════════════════════════════
   ANÁLISIS DE EGRESOS
   ══════════════════════════════════════════════════════════════ */
function readAnalisisEgresos() {
  try {
    var anioActual = new Date().getFullYear();
    var anioAnterior = anioActual - 1;

    var porMes = {};
    var porAnio = {};
    var curContable = { Costo: 0, Gasto: 0, 'Crédito': 0 };
    var curTipo = {};
    var curSub = {};
    var prevSub = {};
    var curProv = {};
    var curTotal = 0, prevTotal = 0;

    Object.keys(EGRESOS_IDS).forEach(function (anio) {
      try {
        var ss = SpreadsheetApp.openById(EGRESOS_IDS[anio]);
        var sh = ss.getSheetByName(EGRESOS_TABS[anio]) || ss.getSheets()[0];
        if (!sh) return;
        var raw = sh.getDataRange().getValues();
        if (raw.length < 2) return;
        var H = raw[0].map(function (h) { return String(h).trim().toLowerCase(); });
        function col(kw) { for (var c = 0; c < H.length; c++) if (H[c].indexOf(kw) > -1) return c; return -1; }
        var iF = col('fecha'), iV = col('vencimiento'), iE = col('egresos'),
            iC = col('contable'), iT = col('tipo'), iS = col('subtipo'), iP = col('proveedor');

        for (var r = 1; r < raw.length; r++) {
          var row = raw[r];
          var f = (iF > -1 ? row[iF] : '') || (iV > -1 ? row[iV] : '');
          var d = (f instanceof Date) ? f : new Date(f);
          if (!d || isNaN(d.getTime())) continue;
          var monto = _anNum(iE > -1 ? row[iE] : 0);
          if (!monto) continue;
          var y = d.getFullYear();
          var mk = y + '-' + String(d.getMonth() + 1).padStart(2, '0');
          var cont = (iC > -1 ? String(row[iC] || '').trim() : '') || 'Gasto';
          var tip = (iT > -1 ? String(row[iT] || '').trim() : '') || 'Variable';
          var sub = (iS > -1 ? String(row[iS] || '').trim() : '') || 'Otros';
          var prov = (iP > -1 ? String(row[iP] || '').trim() : '') || 'Sin proveedor';

          if (!porMes[mk]) porMes[mk] = { costo: 0, gasto: 0, credito: 0, total: 0 };
          var bucket = /costo/i.test(cont) ? 'costo' : (/cr[ée]dito/i.test(cont) ? 'credito' : 'gasto');
          porMes[mk][bucket] += monto; porMes[mk].total += monto;
          porAnio[y] = (porAnio[y] || 0) + monto;

          if (y === anioActual) {
            if (curContable[cont] == null) curContable[cont] = 0;
            curContable[cont] += monto;
            curTipo[tip] = (curTipo[tip] || 0) + monto;
            if (!curSub[sub]) curSub[sub] = { total: 0, contable: cont };
            curSub[sub].total += monto;
            curProv[prov] = (curProv[prov] || 0) + monto;
            curTotal += monto;
          } else if (y === anioAnterior) {
            prevSub[sub] = (prevSub[sub] || 0) + monto;
            prevTotal += monto;
          }
        }
      } catch (e) {}
    });

    var mesesArr = Object.keys(porMes).sort().map(function (mk) {
      return { mes: mk, costo: porMes[mk].costo, gasto: porMes[mk].gasto, credito: porMes[mk].credito, total: porMes[mk].total };
    });

    var subArr = Object.keys(curSub).map(function (s) {
      var tot = curSub[s].total, ant = prevSub[s] || 0;
      return { subtipo: s, contable: curSub[s].contable, total: tot, pct: curTotal > 0 ? (tot / curTotal) * 100 : 0, totalAnt: ant, crec: ant > 0 ? ((tot - ant) / ant * 100) : null, delta: tot - ant };
    }).sort(function (a, b) { return b.total - a.total; });

    var provArr = Object.keys(curProv).map(function (p) {
      return { proveedor: p, total: curProv[p], pct: curTotal > 0 ? (curProv[p] / curTotal) * 100 : 0 };
    }).sort(function (a, b) { return b.total - a.total; }).slice(0, 12);

    var fijo = curTipo['Fijo'] || curTipo['fijo'] || 0;
    var variable = curTotal - fijo;

    var insights = [];
    var crecTotal = prevTotal > 0 ? ((curTotal - prevTotal) / prevTotal) * 100 : null;
    if (subArr.length) {
      var top = subArr[0];
      insights.push({ tipo: top.pct > 25 ? 'warn' : 'info', icono: 'pie-chart', titulo: top.subtipo + ' concentra ' + top.pct.toFixed(0) + '% de tus egresos', detalle: _anFM(top.total) + ' en el año. ' + (top.pct > 25 ? 'Alta concentración: una mejora aquí mueve mucho la aguja.' : 'Es tu categoría más grande.') });
    }
    var crecientes = subArr.filter(function (s) { return s.crec != null && s.delta > 0; }).sort(function (a, b) { return b.delta - a.delta; });
    if (crecientes.length && crecientes[0].crec > 10) {
      var g = crecientes[0];
      insights.push({ tipo: 'warn', icono: 'trending-up', titulo: g.subtipo + ' subió +' + g.crec.toFixed(0) + '% vs el año pasado', detalle: '+' + _anFM(g.delta) + ' más que en ' + anioAnterior + '. El mayor aumento — vale la pena revisar y renegociar.' });
    }
    if (curTotal > 0) {
      var pctFijo = (fijo / curTotal) * 100;
      insights.push({ tipo: pctFijo > 70 ? 'warn' : 'ok', icono: 'lock', titulo: pctFijo.toFixed(0) + '% de tus egresos son fijos', detalle: pctFijo > 70 ? 'Poca flexibilidad para recortar rápido: enfócate en renegociar contratos (renta, software, nómina).' : 'Tienes margen variable para ajustar gasto según la demanda.' });
      var pctCosto = ((curContable['Costo'] || 0) / curTotal) * 100;
      insights.push({ tipo: 'info', icono: 'layers', titulo: pctCosto.toFixed(0) + '% costos directos · ' + (100 - pctCosto).toFixed(0) + '% gastos', detalle: 'Los costos suben con más pacientes (variable sano); los gastos fijos no — ahí está el ahorro estructural.' });
    }
    var candidatos = subArr.filter(function (s) { return s.crec != null && s.crec > 25 && s.pct > 3; }).slice(0, 3);
    if (candidatos.length) {
      insights.push({ tipo: 'warn', icono: 'scissors', titulo: 'Candidatos a revisar para ahorrar', detalle: candidatos.map(function (c) { return c.subtipo + ' (+' + c.crec.toFixed(0) + '%)'; }).join(' · ') + '. Crecieron fuerte y pesan en el total.' });
    }
    if (provArr.length && provArr[0].pct > 20) {
      insights.push({ tipo: 'info', icono: 'briefcase', titulo: provArr[0].proveedor + ' es ' + provArr[0].pct.toFixed(0) + '% de tu gasto', detalle: 'Alta dependencia de un proveedor — buen punto para negociar volumen o buscar alternativa.' });
    }

    return {
      ok: true, anioActual: anioActual, anioAnterior: anioAnterior,
      totalActual: curTotal, totalAnterior: prevTotal, crecTotal: crecTotal,
      porAnio: porAnio, contable: curContable, fijo: fijo, variable: variable,
      meses: mesesArr, subtipos: subArr, proveedores: provArr, insights: insights
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ══════════════════════════════════════════════════════════════
   ANÁLISIS DE PACIENTES
   Lee de PAC_SS_ID la hoja "Pacientes"
   ══════════════════════════════════════════════════════════════ */
function readAnalisisPacientes() {
  try {
    var ssPac = SpreadsheetApp.openById(PAC_SS_ID);
    var sh = ssPac.getSheetByName('Pacientes') || ssPac.getSheets()[0];
    if (!sh) return { ok: false, error: 'No se encontró la hoja de Pacientes' };
    var raw = sh.getDataRange().getValues();
    if (raw.length < 2) return { ok: true, total: 0, rows: [], insights: [] };

    var H = raw[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
    function col(kw) { for (var c=0;c<H.length;c++) if (H[c].indexOf(kw)>-1) return c; return -1; }
    var iFecha = col('alta') > -1 ? col('alta') : col('fecha');
    var iNombre = col('nombre') > -1 ? col('nombre') : 0;
    var iLista = col('lista');
    var iEstado = col('estado') > -1 ? col('estado') : col('estatus') > -1 ? col('estatus') : -1;

    var anioActual = new Date().getFullYear();
    var porMes = {}, porAnio = {}, porLista = {}, porEstado = {};
    var total = 0, nuevosActual = 0, nuevosAnterior = 0;

    for (var r = 1; r < raw.length; r++) {
      var row = raw[r];
      var nombre = String(row[iNombre]||'').trim(); if (!nombre) continue;
      total++;

      if (iFecha > -1 && row[iFecha]) {
        var f = row[iFecha]; var d = (f instanceof Date) ? f : new Date(f);
        if (d && !isNaN(d.getTime())) {
          var y = d.getFullYear(), mo = d.getMonth()+1;
          var mk = y+'-'+String(mo).padStart(2,'0');
          porMes[mk] = (porMes[mk]||0)+1;
          porAnio[y] = (porAnio[y]||0)+1;
          if (y === anioActual) nuevosActual++;
          if (y === anioActual-1) nuevosAnterior++;
        }
      }

      if (iLista > -1) {
        var lista = String(row[iLista]||'').trim()||'Sin lista';
        porLista[lista] = (porLista[lista]||0)+1;
      }
      if (iEstado > -1) {
        var estado = String(row[iEstado]||'').trim()||'Sin estado';
        porEstado[estado] = (porEstado[estado]||0)+1;
      }
    }

    var mesesArr = Object.keys(porMes).sort().slice(-18).map(function(mk){ return { mes:mk, count:porMes[mk] }; });

    // Tendencia: promedio nuevos/mes últimos 3 vs 3 anteriores
    var recent3 = mesesArr.slice(-3).reduce(function(s,x){ return s+x.count; },0) / 3;
    var prev3 = mesesArr.slice(-6,-3).reduce(function(s,x){ return s+x.count; },0) / 3;
    var tendenciaPac = prev3 > 0 ? (recent3 - prev3) / prev3 : null;

    var crecAnual = nuevosAnterior > 0 ? (nuevosActual - nuevosAnterior) / nuevosAnterior : null;

    var listaArr = Object.keys(porLista).map(function(l){ return { lista:l, count:porLista[l], pct:total>0?porLista[l]/total*100:0 }; }).sort(function(a,b){ return b.count-a.count; });
    var estadoArr = Object.keys(porEstado).map(function(e){ return { estado:e, count:porEstado[e] }; }).sort(function(a,b){ return b.count-a.count; });

    var insights = [];
    if (total > 0) insights.push({ tipo:'info', icono:'users', titulo:total+' pacientes en la base de datos', detalle:'Nuevos '+anioActual+': '+nuevosActual+(crecAnual!==null?' · Crecimiento: '+(crecAnual>=0?'+':'')+( crecAnual*100).toFixed(0)+'% YoY':'') });
    if (tendenciaPac!==null) {
      insights.push({ tipo:tendenciaPac>=0?'ok':'warn', icono:tendenciaPac>=0?'trending-up':'trending-down',
        titulo:'Adquisición reciente: '+(tendenciaPac>=0?'+':'')+( tendenciaPac*100).toFixed(0)+'% vs período anterior',
        detalle:'Promedio últimos 3 meses: '+recent3.toFixed(1)+' nuevos/mes vs '+prev3.toFixed(1)+' previos.' });
    }
    if (listaArr.length) {
      insights.push({ tipo:'info', icono:'layers', titulo:'Segmentación: '+listaArr[0].lista+' es '+listaArr[0].pct.toFixed(0)+'% del total', detalle:listaArr.slice(0,3).map(function(l){ return l.lista+': '+l.count; }).join(' · ') });
    }

    return {
      ok:true, total:total, nuevosActual:nuevosActual, nuevosAnterior:nuevosAnterior,
      crecAnual:crecAnual, tendencia:tendenciaPac, avgNuevosMes:recent3,
      meses:mesesArr, porAnio:porAnio, listas:listaArr, estados:estadoArr, insights:insights
    };
  } catch(ex) { return { ok:false, error:ex.message }; }
}

/* ══════════════════════════════════════════════════════════════
   ANÁLISIS DE SERVICIOS
   Mix de ingresos por línea de servicio: ER (autoritativo) + BD_Ingresos (operativo)
   ══════════════════════════════════════════════════════════════ */
function readAnalisisServicios() {
  try {
    var anioActual = new Date().getFullYear();
    var eb = _anLoadErBudget();
    var erByLine = eb.er.byLine;   // {'Surrogacy': {'2026-04':200000}, ...}
    var bgByLine = eb.budgetByLine || {};   // budget por línea (sin reabrir el ER)
    var bd = _anReadBDIngresos();

    // Totales anuales por línea desde ER (autoritativo)
    var erLineas = {};
    Object.keys(erByLine).forEach(function(linea) {
      var data = erByLine[linea];
      var total = 0, totalAnt = 0;
      Object.keys(data).forEach(function(ym) {
        if (ym.substring(0,4) == anioActual) total += data[ym]||0;
        if (ym.substring(0,4) == (anioActual-1)) totalAnt += data[ym]||0;
      });
      var budgetTotal = 0;
      if (bgByLine[linea]) Object.keys(bgByLine[linea]).forEach(function(ym) {
        if (ym.substring(0,4) == anioActual) budgetTotal += bgByLine[linea][ym]||0;
      });
      erLineas[linea] = { linea:linea, total:total, totalAnt:totalAnt, budget:budgetTotal,
        crec:totalAnt>0?(total-totalAnt)/totalAnt:null, pctBudget:budgetTotal>0?total/budgetTotal:null };
    });

    var totalER = Object.keys(erLineas).reduce(function(s,l){ return s+erLineas[l].total; },0);
    var lineasArr = Object.keys(erLineas).map(function(l) {
      var k=erLineas[l]; return Object.assign({}, k, { pct: totalER>0?k.total/totalER*100:0 });
    }).sort(function(a,b){ return b.total-a.total; });

    // Mix operativo desde BD_Ingresos (categorías → operaciones, ticket)
    var opMix = [];
    if (bd) {
      var byCatActual = {};
      Object.keys(bd.byCatYear).forEach(function(key) {
        var k=bd.byCatYear[key]; if (k.y===anioActual) byCatActual[k.cat]=k;
      });
      var totalOp = Object.keys(byCatActual).reduce(function(s,c){ return s+byCatActual[c].total; },0);
      opMix = Object.keys(byCatActual).map(function(c) {
        var k=byCatActual[c]; var ops=Object.keys(k.ops).length;
        return { categoria:c, total:k.total, pct:totalOp>0?k.total/totalOp*100:0, operaciones:ops, ticket:ops>0?k.total/ops:0, cantidad:k.cant };
      }).sort(function(a,b){ return b.total-a.total; });
    }

    // Serie mensual por línea (últimos 12 meses)
    var hoy = new Date();
    var meses12 = [];
    for (var i=11;i>=0;i--) {
      var d=new Date(hoy.getFullYear(), hoy.getMonth()-i, 1);
      meses12.push(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'));
    }
    var serieLineas = Object.keys(erByLine).slice(0,6).map(function(linea) {
      return { linea:linea, data: meses12.map(function(ym){ return erByLine[linea][ym]||0; }) };
    });

    var insights = [];
    if (lineasArr.length) {
      insights.push({ tipo:'ok', icono:'star', titulo:lineasArr[0].linea+' lidera con '+lineasArr[0].pct.toFixed(0)+'% del mix', detalle:'Ingresos '+anioActual+': '+_anFM(lineasArr[0].total)+(lineasArr[0].crec!==null?' · '+(lineasArr[0].crec>=0?'+':'')+( lineasArr[0].crec*100).toFixed(0)+'% YoY':'') });
    }
    lineasArr.filter(function(l){ return l.crec!==null&&l.crec>20; }).slice(0,1).forEach(function(l) {
      insights.push({ tipo:'ok', icono:'trending-up', titulo:l.linea+' creció +'+(l.crec*100).toFixed(0)+'% vs año anterior', detalle:'De '+_anFM(l.totalAnt)+' a '+_anFM(l.total)+' — línea en expansión.' });
    });
    lineasArr.filter(function(l){ return l.pctBudget!==null&&l.pctBudget<0.9; }).slice(0,1).forEach(function(l) {
      insights.push({ tipo:'warn', icono:'alert-triangle', titulo:l.linea+' al '+(l.pctBudget*100).toFixed(0)+'% del budget', detalle:'Real: '+_anFM(l.total)+' vs Plan: '+_anFM(l.budget)+' · Brecha: '+_anFM(l.budget-l.total) });
    });

    return {
      ok:true, anioActual:anioActual, fuente:'Estado de Resultados + BD_Ingresos',
      totalER:totalER, lineas:lineasArr, meses12:meses12,
      serieLineas:serieLineas, opMix:opMix, insights:insights
    };
  } catch(ex) { return { ok:false, error:ex.message }; }
}

/* ══════════════════════════════════════════════════════════════
   ANÁLISIS DE SURROGACY
   Filtrado de BD_Ingresos + línea ER Surrogacy
   ══════════════════════════════════════════════════════════════ */
function readAnalisisSurrogacy() {
  try {
    var anioActual = new Date().getFullYear();
    var SURROGACY_CATS = ['surrogacy','subrogac','maternidad subrogada','surrogate','surrogacy-gm','surrogacy gm'];

    function isSurrogacyCat(cat) {
      var cl = cat.toLowerCase();
      for (var i=0;i<SURROGACY_CATS.length;i++) if (cl.indexOf(SURROGACY_CATS[i])>-1) return true;
      return false;
    }

    // ER línea Surrogacy
    var eb = _anLoadErBudget();
    var erSurr = {}, bgSurr = {};
    Object.keys(eb.er.byLine).forEach(function(l) {
      if (isSurrogacyCat(l)) Object.keys(eb.er.byLine[l]).forEach(function(ym) {
        erSurr[ym] = (erSurr[ym]||0) + (eb.er.byLine[l][ym]||0);
      });
    });

    // BD_Ingresos filtrado Surrogacy
    var bd = _anReadBDIngresos();
    var opsSurr = {}, totalSurr = 0, totalSurrAnt = 0;
    var prodSurr = {}, porMes = {};
    if (bd) {
      Object.keys(bd.byCatYear).forEach(function(key) {
        var k=bd.byCatYear[key];
        if (!isSurrogacyCat(k.cat)) return;
        if (k.y===anioActual) { totalSurr+=k.total; }
        if (k.y===anioActual-1) { totalSurrAnt+=k.total; }
      });
      Object.keys(bd.byProd).forEach(function(p) {
        var k=bd.byProd[p]; if (!isSurrogacyCat(k.cat)) return;
        var ops=Object.keys(k.ops).length;
        prodSurr[p]={ prod:p, cat:k.cat, total:k.total, ops:ops, ticket:ops>0?k.total/ops:0, cant:k.cant };
      });
      Object.keys(bd.opSet).forEach(function(op) {
        var k=bd.opSet[op]; if (!isSurrogacyCat(k.cat)) return;
        opsSurr[op]=k;
      });
    }

    var totalSurrER = Object.keys(erSurr).filter(function(ym){ return ym.substring(0,4)==anioActual; }).reduce(function(s,ym){ return s+(erSurr[ym]||0); },0);
    var numCasos = Object.keys(opsSurr).length;
    var ticketProm = numCasos>0 ? totalSurr/numCasos : 0;
    var crecYoY = totalSurrAnt>0 ? (totalSurr-totalSurrAnt)/totalSurrAnt : null;

    var mesesArr = Object.keys(erSurr).sort().slice(-18).map(function(ym) {
      return { mes:ym, total:erSurr[ym]||0 };
    });

    var prodArr = Object.keys(prodSurr).map(function(p){ return prodSurr[p]; }).sort(function(a,b){ return b.total-a.total; });

    var insights = [];
    insights.push({ tipo:'info', icono:'heart', titulo:numCasos+' casos Surrogacy registrados en '+anioActual, detalle:'Ingresos totales (operativo): '+_anFM(totalSurr)+(ticketProm>0?' · Ticket promedio: '+_anFM(ticketProm):'') });
    if (crecYoY!==null) insights.push({ tipo:crecYoY>=0?'ok':'warn', icono:'trending-up', titulo:(crecYoY>=0?'+':'')+( crecYoY*100).toFixed(0)+'% vs '+( anioActual-1), detalle:'Surrogacy: '+_anFM(totalSurr)+' vs '+_anFM(totalSurrAnt)+' el año anterior.' });
    if (totalSurrER>0) insights.push({ tipo:'ok', icono:'bar-chart', titulo:'Ingresos ER Surrogacy: '+_anFM(totalSurrER), detalle:'Fuente Estado de Resultados (datos contables del '+anioActual+').' });

    return {
      ok:true, anioActual:anioActual, fuente:'BD_Ingresos + ER',
      totalBD:totalSurr, totalER:totalSurrER, totalAnt:totalSurrAnt,
      numCasos:numCasos, ticketPromedio:ticketProm, crecYoY:crecYoY,
      meses:mesesArr, productos:prodArr, insights:insights
    };
  } catch(ex) { return { ok:false, error:ex.message }; }
}

/* ══════════════════════════════════════════════════════════════
   ANÁLISIS DE RENTABILIDAD
   Combina: ER (ingresos) + Egresos (gastos) → P&L mensual
   ══════════════════════════════════════════════════════════════ */
function readAnalisisRentabilidad() {
  try {
    var anioActual = new Date().getFullYear();
    var hoy = new Date();
    var mesActual = anioActual+'-'+String(hoy.getMonth()+1).padStart(2,'0');

    // 1. ER: ingresos y budget
    var eb = _anLoadErBudget();
    var erIncome = eb.er.income;
    var budget = eb.budget;

    // 2. Egresos mensuales (desde readAnalisisEgresos ya calculado)
    var egPorMes = {};
    Object.keys(EGRESOS_IDS).forEach(function(anio) {
      try {
        var ss = SpreadsheetApp.openById(EGRESOS_IDS[anio]);
        var sh = ss.getSheetByName(EGRESOS_TABS[anio]) || ss.getSheets()[0];
        if (!sh) return;
        var raw = sh.getDataRange().getValues();
        if (raw.length < 2) return;
        var H = raw[0].map(function(h){ return String(h||'').toLowerCase(); });
        function col(kw){ for(var c=0;c<H.length;c++) if(H[c].indexOf(kw)>-1) return c; return -1; }
        var iF=col('fecha'), iV=col('vencimiento'), iE=col('egresos');
        for (var r=1;r<raw.length;r++) {
          var row=raw[r];
          var f=(iF>-1?row[iF]:'') || (iV>-1?row[iV]:'');
          var d=(f instanceof Date)?f:new Date(f);
          if (!d||isNaN(d.getTime())) continue;
          var monto=_anNum(iE>-1?row[iE]:0); if (!monto) continue;
          var mk=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
          egPorMes[mk]=(egPorMes[mk]||0)+monto;
        }
      } catch(e){}
    });

    // 3. Serie mensual P&L
    var allYms = {};
    Object.keys(erIncome).forEach(function(ym){ allYms[ym]=1; });
    Object.keys(egPorMes).forEach(function(ym){ allYms[ym]=1; });
    var mesesArr = Object.keys(allYms).sort().map(function(ym) {
      var inc=erIncome[ym]||0, eg=egPorMes[ym]||0, bud=budget[ym]||0;
      return { mes:ym, ingresos:inc, egresos:eg, margen:inc-eg, margenPct:inc>0?(inc-eg)/inc*100:0, budget:bud };
    });

    // 4. YTD KPIs
    var ytdInc=0, ytdEg=0, ytdBud=0, ytdIncAnt=0, ytdEgAnt=0;
    mesesArr.forEach(function(m) {
      var y=m.mes.substring(0,4), mo=m.mes.substring(5);
      if (y==anioActual && m.mes<=mesActual) { ytdInc+=m.ingresos; ytdEg+=m.egresos; ytdBud+=m.budget; }
      if (y==(anioActual-1) && mo<=mesActual.substring(5)) { ytdIncAnt+=m.ingresos; ytdEgAnt+=m.egresos; }
    });
    var ytdMargen = ytdInc - ytdEg;
    var ytdMargenPct = ytdInc > 0 ? ytdMargen / ytdInc * 100 : 0;
    var ytdEBITDA = ytdMargen; // simplificado sin D&A
    var margenAnt = ytdIncAnt - ytdEgAnt;
    var crecMargen = margenAnt !== 0 ? (ytdMargen - margenAnt) / Math.abs(margenAnt) : null;

    // 5. Mejor y peor mes del año actual
    var mesesAnio = mesesArr.filter(function(m){ return m.mes.substring(0,4)==anioActual && m.mes<=mesActual && m.ingresos>0; });
    mesesAnio.sort(function(a,b){ return b.margenPct - a.margenPct; });
    var mejorMes = mesesAnio[0] || null;
    var peorMes = mesesAnio[mesesAnio.length-1] || null;

    var insights = [];
    insights.push({ tipo: ytdMargenPct>=30?'ok':(ytdMargenPct>=15?'info':'warn'), icono:'percent',
      titulo: 'Margen operativo YTD: '+ytdMargenPct.toFixed(1)+'%',
      detalle: 'Ingresos: '+_anFM(ytdInc)+' · Egresos: '+_anFM(ytdEg)+' · Margen: '+_anFM(ytdMargen) });
    if (ytdBud>0) {
      var pctBudget = ytdInc/ytdBud;
      insights.push({ tipo:pctBudget>=1?'ok':'warn', icono:'target', titulo:'Budget: al '+(pctBudget*100).toFixed(0)+'% del plan',
        detalle:'Real '+_anFM(ytdInc)+' vs Plan '+_anFM(ytdBud)+' · Diferencia: '+_anFM(ytdInc-ytdBud) });
    }
    if (crecMargen!==null) {
      insights.push({ tipo:crecMargen>=0?'ok':'warn', icono:crecMargen>=0?'trending-up':'trending-down',
        titulo:'Margen '+(crecMargen>=0?'mejoró':'empeoró')+' '+(crecMargen>=0?'+':'')+( crecMargen*100).toFixed(0)+'% YoY',
        detalle:'Margen YTD '+anioActual+': '+_anFM(ytdMargen)+' vs '+_anFM(margenAnt)+' mismo período '+( anioActual-1) });
    }
    if (mejorMes) insights.push({ tipo:'ok', icono:'award', titulo:'Mejor mes del año: '+mejorMes.mes+' ('+mejorMes.margenPct.toFixed(1)+'%)', detalle:'Ingresos: '+_anFM(mejorMes.ingresos)+' · Egresos: '+_anFM(mejorMes.egresos)+' · Margen: '+_anFM(mejorMes.margen) });

    return {
      ok:true, anioActual:anioActual, fuente:'Estado de Resultados + Egresos',
      ytdIngresos:ytdInc, ytdEgresos:ytdEg, ytdMargen:ytdMargen, ytdMargenPct:ytdMargenPct,
      ytdBudget:ytdBud, ytdEBITDA:ytdEBITDA, crecMargen:crecMargen,
      meses:mesesArr, insights:insights
    };
  } catch(ex) { return { ok:false, error:ex.message }; }
}

/* ==============================================================
   ESTADO DE CUENTA POR PACIENTE
   Lee BD_Ingresos de todos los años disponibles y filtra por nombre
   ============================================================== */

function readEstadoCuentaPaciente(pacienteNombre) {
  try {
    if (!pacienteNombre) return { ok: false, error: 'Nombre de paciente requerido' };
    var nombreBuscar = String(pacienteNombre).trim().toLowerCase();

    // BD_Ingresos es la hoja CONSOLIDADA de todos los años; vive solo en
    // INGRESOS_SS_ID. No leer de los spreadsheets por año (2024/2025)
    // porque causaría duplicados si esas hojas también tienen BD_Ingresos.
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() === BD_INGRESOS_TAB) { sh = sheets[i]; break; }
    }
    if (!sh) return { ok: false, error: 'No se encontró la hoja ' + BD_INGRESOS_TAB + ' en el spreadsheet principal.' };

    var data = sh.getDataRange().getValues();
    var H = data[0].map(function(x){ return String(x||'').trim().toLowerCase(); });
    function hc(){ for (var a=0;a<arguments.length;a++){ var k=H.indexOf(arguments[a]); if(k>-1) return k; } return -1; }
    var iOp    = hc('op');
    var iFecha = hc('fecha');
    var iPac   = hc('paciente');
    var iCat   = hc('categoria','categoría');
    var iProd  = hc('producto');
    var iDesc  = hc('descripcion','descripción');
    var iCant  = hc('cantidad');
    var iTotal = hc('totalpagar','total a pagar','total');
    var iPagado= hc('pagado');
    var iEst   = hc('estatus','estado');
    var iPago  = hc('formapago','forma de pago','forma pago');
    if (iPac < 0 || iTotal < 0)
      return { ok: false, error: 'Columnas Paciente o Total no encontradas en ' + BD_INGRESOS_TAB };

    var movimientos = [];
    var totalGeneral = 0;
    // Total efectivamente PAGADO por el paciente. Convención consistente con el
    // histórico: si la columna Pagado viene en blanco/0 se asume la línea pagada por
    // completo (así lo trata también el cálculo de saldo por línea de abajo). Si
    // Pagado trae un número se toma literal — puede ser MENOR que el total (pago
    // parcial → nos debe) o MAYOR (p. ej. un ingreso editado a un producto más
    // barato → saldo a favor del paciente).
    var totalPagado = 0;
    // Saldo por pago PARCIAL a nivel línea (solo si Pagado>0 y Pagado<Total; los
    // históricos con Pagado en blanco se asumen pagados). Se agrupa por OP para no
    // duplicar contra el registro explícito de Cuentas por Cobrar más abajo.
    var _saldoLineaPorOp = {};

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var pac = String(row[iPac] || '').trim().toLowerCase();
      if (pac !== nombreBuscar) continue;
      var rawFecha = row[iFecha];
      var fechaStr = '';
      if (rawFecha instanceof Date) {
        fechaStr = Utilities.formatDate(rawFecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        fechaStr = String(rawFecha || '').substring(0, 10);
      }
      var total = parseFloat(String(row[iTotal] || '0').replace(/[$,]/g, '')) || 0;
      var pagadoRow = iPagado >= 0 ? (parseFloat(String(row[iPagado] || '0').replace(/[$,]/g, '')) || 0) : 0;
      var saldoRow = (iPagado >= 0 && pagadoRow > 0.01 && pagadoRow < total - 0.01) ? (total - pagadoRow) : 0;
      var opRow = String(row[iOp] || '').trim();
      if (saldoRow > 0.01) _saldoLineaPorOp[opRow] = (_saldoLineaPorOp[opRow] || 0) + saldoRow;
      // Pagado efectivo: Pagado en blanco/0 ⇒ línea asumida pagada por completo;
      // un número ⇒ se respeta (puede exceder el total = sobrepago / saldo a favor).
      var pagEfectivo = (iPagado < 0) ? total : (pagadoRow > 0.01 ? pagadoRow : total);
      totalPagado += pagEfectivo;
      totalGeneral += total;
      movimientos.push({
        op:       opRow,
        fecha:    fechaStr,
        anio:     fechaStr ? fechaStr.substring(0, 4) : 'Sin año',
        cat:      String(row[iCat]  || '').trim(),
        producto: String(row[iProd] || '').trim(),
        desc:     iDesc >= 0 ? String(row[iDesc] || '').trim() : '',
        cant:     iCant >= 0 ? (parseFloat(String(row[iCant]||'1'))||1) : 1,
        total:    total,
        pagado:   pagadoRow,
        saldo:    saldoRow,
        estatus:  iEst  >= 0 ? String(row[iEst]  || '').trim() : '',
        pago:     iPago >= 0 ? String(row[iPago]  || '').trim() : ''
      });
    }

    // ── Saldo pendiente / "Nos debe" ────────────────────────────────────────────
    // Fuente (b): registro EXPLÍCITO de Cuentas por Cobrar (Motor A de cobranza.gs)
    //   — incluye los adeudos auto-generados al capturar/editar un ingreso con pago
    //   parcial y los saldos iniciales/cargos a crédito. Es la fuente canónica y la
    //   misma que ve el módulo "Cuentas por Cobrar" (evita la deuda falsa por precios
    //   viejos que daría restar Pagado a ciegas de TODO el histórico).
    // Fuente (a): pagos parciales a nivel línea de BD_Ingresos que AÚN no tengan un
    //   renglón en el registro (históricos previos a esta función) — se suman sin
    //   duplicar los que ya están cubiertos por el registro.
    var deudaDetalle = [];
    var saldoPendiente = 0;
    var _opsCubiertos = {};
    try {
      if (typeof _cobReadCargos === 'function' && typeof _cobReadAbonos === 'function' && typeof _cobKeyNom === 'function') {
        var _keyPac = _cobKeyNom(pacienteNombre);
        var _cargos = _cobReadCargos();
        var _abonos = _cobReadAbonos();
        for (var cgi = 0; cgi < _cargos.rows.length; cgi++) {
          var cg = _cargos.rows[cgi];
          if (_cobKeyNom(cg.paciente) !== _keyPac) continue;
          var st = String(cg.estatus || '').toLowerCase();
          if (st === 'cancelado' || st === 'pagado') continue;
          var abo = cg.op ? (_abonos.byOp[cg.op] || 0) : 0;
          var sReg = cg.monto - abo;
          if (sReg <= 0.01) continue;
          if (cg.op) _opsCubiertos[cg.op] = true;
          saldoPendiente += sReg;
          deudaDetalle.push({ op: cg.op || '—', concepto: cg.concepto || 'Saldo', fecha: cg.fecha,
                              saldo: sReg, origen: (String(cg.nota || '').toLowerCase().indexOf('auto-ingreso') > -1) ? 'ingreso' : 'registro' });
        }
      }
    } catch (eCob) {}
    // Pagos parciales de línea no cubiertos por el registro explícito.
    for (var opk in _saldoLineaPorOp) {
      if (_opsCubiertos[opk]) continue;
      saldoPendiente += _saldoLineaPorOp[opk];
      deudaDetalle.push({ op: opk || '—', concepto: 'Pago parcial de OP ' + opk, fecha: '',
                          saldo: _saldoLineaPorOp[opk], origen: 'ingreso' });
    }
    deudaDetalle.sort(function (a, b) { return b.saldo - a.saldo; });

    movimientos.sort(function(a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });

    var porCategoria = {};
    movimientos.forEach(function(m) {
      porCategoria[m.cat || 'Sin categoría'] = (porCategoria[m.cat || 'Sin categoría'] || 0) + m.total;
    });

    var porAnio = {};
    movimientos.forEach(function(m) {
      porAnio[m.anio] = (porAnio[m.anio] || 0) + m.total;
    });

    // Saldo NETO (con signo) = Facturado − Pagado. Positivo = nos debe;
    // negativo = saldo a favor del paciente (sobrepago). totalFacturado es alias
    // explícito de totalGeneral (el "Total histórico" que ya se mostraba).
    var totalFacturado = totalGeneral;
    var saldoNeto = totalFacturado - totalPagado;

    return {
      ok: true,
      paciente: String(pacienteNombre).trim(),
      totalGeneral: totalGeneral,
      totalFacturado: totalFacturado,
      totalPagado: totalPagado,
      saldoNeto: saldoNeto,
      totalMovimientos: movimientos.length,
      movimientos: movimientos,
      porCategoria: porCategoria,
      porAnio: porAnio,
      saldoPendiente: saldoPendiente,
      deudaDetalle: deudaDetalle
    };
  } catch(ex) {
    return { ok: false, error: ex.message };
  }
}
