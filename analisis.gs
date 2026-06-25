/* ==============================================================
   analisis.gs — Centro de Análisis de Egresos
   --------------------------------------------------------------
   Lee TODO el histórico de egresos (EGRESOS_IDS por año) y genera
   un análisis para detectar dónde se gasta más, qué creció y dónde
   se puede recortar. Incluye recomendaciones automáticas.

   Ruta:  GET ?action=analisisEgresos
   ============================================================== */

function _anNum(v) { if (typeof v === 'number') return v; var n = parseFloat(String(v || '').replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
function _anFM(v) { v = Math.abs(Number(v) || 0); if (v >= 1e6) return '$' + (v / 1e6).toFixed(2) + 'M'; if (v >= 1e3) return '$' + Math.round(v / 1e3) + 'K'; return '$' + Math.round(v); }

function readAnalisisIngresos() {
  try {
    var anioActual = new Date().getFullYear();
    var anioAnterior = anioActual - 1;
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = null, all = ss.getSheets();
    for (var i = 0; i < all.length; i++) if (all[i].getName() === BD_INGRESOS_TAB) { sh = all[i]; break; }
    if (!sh) return { ok: false, error: 'No existe ' + BD_INGRESOS_TAB };
    var raw = sh.getDataRange().getValues();

    var porMes = {}, porAnio = {};
    var curCat = {}, prevCat = {}, curProd = {}, curOps = {};
    var curTotal = 0, prevTotal = 0;

    for (var r = 1; r < raw.length; r++) {
      var row = raw[r];
      var op = String(row[0] || '').trim(); if (!op) continue;
      var f = row[2]; var d = (f instanceof Date) ? f : new Date(f); if (!d || isNaN(d.getTime())) continue;
      var monto = _anNum(row[9]);                 // TotalPagar
      var cant = _anNum(row[8]) || 0;             // Cantidad
      var cat = String(row[4] || '').trim() || 'Sin categoría';
      var prod = String(row[5] || '').trim() || cat;
      var y = d.getFullYear();
      var mk = y + '-' + String(d.getMonth() + 1).padStart(2, '0');
      if (!porMes[mk]) porMes[mk] = { total: 0 };
      porMes[mk].total += monto; porAnio[y] = (porAnio[y] || 0) + monto;

      if (y === anioActual) {
        if (!curCat[cat]) curCat[cat] = { total: 0, cant: 0, ops: {} };
        curCat[cat].total += monto; curCat[cat].cant += cant; curCat[cat].ops[op] = 1;
        if (!curProd[prod]) curProd[prod] = { total: 0, cant: 0, categoria: cat };
        curProd[prod].total += monto; curProd[prod].cant += cant;
        curTotal += monto; curOps[op] = 1;
      } else if (y === anioAnterior) {
        prevCat[cat] = (prevCat[cat] || 0) + monto; prevTotal += monto;
      }
    }

    var categorias = Object.keys(curCat).map(function (c) {
      var t = curCat[c].total, ant = prevCat[c] || 0, ops = Object.keys(curCat[c].ops).length;
      return { categoria: c, total: t, pct: curTotal > 0 ? t / curTotal * 100 : 0, cantidad: curCat[c].cant, operaciones: ops, ticket: ops > 0 ? t / ops : 0, totalAnt: ant, crec: ant > 0 ? ((t - ant) / ant * 100) : null };
    }).sort(function (a, b) { return b.total - a.total; });

    var productos = Object.keys(curProd).map(function (p) {
      return { producto: p, categoria: curProd[p].categoria, total: curProd[p].total, cantidad: curProd[p].cant, ticket: curProd[p].cant > 0 ? curProd[p].total / curProd[p].cant : 0 };
    });
    var topRevenue = productos.slice().sort(function (a, b) { return b.total - a.total; }).slice(0, 12);
    var masVendidos = productos.slice().filter(function (p) { return p.cantidad > 0; }).sort(function (a, b) { return b.cantidad - a.cantidad; }).slice(0, 12);
    var mesesArr = Object.keys(porMes).sort().map(function (mk) { return { mes: mk, total: porMes[mk].total }; });
    var crecTotal = prevTotal > 0 ? ((curTotal - prevTotal) / prevTotal * 100) : null;
    var numOps = Object.keys(curOps).length;

    var insights = [];
    if (categorias.length) {
      var top = categorias[0];
      insights.push({ tipo: 'ok', icono: 'award', titulo: top.categoria + ' es tu mayor ingreso: ' + top.pct.toFixed(0) + '% del total', detalle: _anFM(top.total) + ' en ' + anioActual + '. Tu motor principal de ingresos.' });
    }
    if (masVendidos.length) {
      var mv = masVendidos[0];
      insights.push({ tipo: 'info', icono: 'package', titulo: 'El más vendido por volumen: ' + mv.producto, detalle: Math.round(mv.cantidad) + ' unidades · ' + _anFM(mv.total) + ' en ingresos.' });
    }
    var crecientes = categorias.filter(function (c) { return c.crec != null; }).sort(function (a, b) { return b.crec - a.crec; });
    if (crecientes.length && crecientes[0].crec > 10) {
      insights.push({ tipo: 'ok', icono: 'trending-up', titulo: crecientes[0].categoria + ' creció +' + crecientes[0].crec.toFixed(0) + '% vs ' + anioAnterior, detalle: 'El servicio con mayor crecimiento — buen candidato para impulsar más.' });
    }
    var cayendo = categorias.filter(function (c) { return c.crec != null && c.crec < -5; }).sort(function (a, b) { return a.crec - b.crec; });
    if (cayendo.length) {
      insights.push({ tipo: 'warn', icono: 'trending-down', titulo: cayendo[0].categoria + ' cayó ' + cayendo[0].crec.toFixed(0) + '% vs el año pasado', detalle: 'Revisa por qué bajó — ¿precio, demanda o competencia?' });
    }
    var conTicket = categorias.filter(function (c) { return c.operaciones >= 3; }).sort(function (a, b) { return b.ticket - a.ticket; });
    if (conTicket.length) {
      insights.push({ tipo: 'info', icono: 'gem', titulo: 'Mejor ticket promedio: ' + conTicket[0].categoria, detalle: _anFM(conTicket[0].ticket) + ' por operación — servicio de alto valor para priorizar en ventas.' });
    }

    return {
      ok: true, anioActual: anioActual, anioAnterior: anioAnterior,
      totalActual: curTotal, totalAnterior: prevTotal, crecTotal: crecTotal,
      numOperaciones: numOps, ticketPromedio: numOps > 0 ? curTotal / numOps : 0,
      porAnio: porAnio, meses: mesesArr, categorias: categorias,
      topRevenue: topRevenue, masVendidos: masVendidos, insights: insights
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function readAnalisisEgresos() {
  try {
    var anioActual = new Date().getFullYear();
    var anioAnterior = anioActual - 1;

    var porMes = {};          // 'YYYY-MM' -> {costo,gasto,credito,total}
    var porAnio = {};         // year -> total
    var curContable = { Costo: 0, Gasto: 0, 'Crédito': 0 };
    var curTipo = {};         // Fijo/Variable -> total (año actual)
    var curSub = {};          // subtipo -> {total, contable}
    var prevSub = {};         // subtipo -> total (año anterior)
    var curProv = {};         // proveedor -> total (año actual)
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

    // ── Series mensual ordenada ──
    var mesesArr = Object.keys(porMes).sort().map(function (mk) {
      return { mes: mk, costo: porMes[mk].costo, gasto: porMes[mk].gasto, credito: porMes[mk].credito, total: porMes[mk].total };
    });

    // ── Ranking de subtipos (año actual) con crecimiento vs anterior ──
    var subArr = Object.keys(curSub).map(function (s) {
      var tot = curSub[s].total, ant = prevSub[s] || 0;
      return {
        subtipo: s, contable: curSub[s].contable, total: tot,
        pct: curTotal > 0 ? (tot / curTotal) * 100 : 0,
        totalAnt: ant, crec: ant > 0 ? ((tot - ant) / ant) * 100 : null, delta: tot - ant
      };
    }).sort(function (a, b) { return b.total - a.total; });

    // ── Top proveedores (año actual) ──
    var provArr = Object.keys(curProv).map(function (p) {
      return { proveedor: p, total: curProv[p], pct: curTotal > 0 ? (curProv[p] / curTotal) * 100 : 0 };
    }).sort(function (a, b) { return b.total - a.total; }).slice(0, 12);

    // ── Fijo vs Variable ──
    var fijo = curTipo['Fijo'] || curTipo['fijo'] || 0;
    var variable = curTotal - fijo;

    // ── Recomendaciones automáticas ──
    var insights = [];
    var crecTotal = prevTotal > 0 ? ((curTotal - prevTotal) / prevTotal) * 100 : null;

    if (subArr.length) {
      var top = subArr[0];
      insights.push({ tipo: top.pct > 25 ? 'warn' : 'info', icono: 'pie-chart',
        titulo: top.subtipo + ' concentra ' + top.pct.toFixed(0) + '% de tus egresos',
        detalle: _anFM(top.total) + ' en el año. ' + (top.pct > 25 ? 'Alta concentración: una mejora aquí mueve mucho la aguja.' : 'Es tu categoría más grande.') });
    }
    // Mayor incremento absoluto vs año anterior
    var crecientes = subArr.filter(function (s) { return s.crec != null && s.delta > 0; }).sort(function (a, b) { return b.delta - a.delta; });
    if (crecientes.length && crecientes[0].crec > 10) {
      var g = crecientes[0];
      insights.push({ tipo: 'warn', icono: 'trending-up',
        titulo: g.subtipo + ' subió +' + g.crec.toFixed(0) + '% vs el año pasado',
        detalle: '+' + _anFM(g.delta) + ' más que en ' + anioAnterior + '. El mayor aumento — vale la pena revisar y renegociar.' });
    }
    // Fijo vs variable
    if (curTotal > 0) {
      var pctFijo = (fijo / curTotal) * 100;
      insights.push({ tipo: pctFijo > 70 ? 'warn' : 'ok', icono: 'lock',
        titulo: pctFijo.toFixed(0) + '% de tus egresos son fijos',
        detalle: pctFijo > 70 ? 'Poca flexibilidad para recortar rápido: enfócate en renegociar contratos (renta, software, nómina).' : 'Tienes margen variable para ajustar gasto según la demanda.' });
    }
    // Costos vs gastos
    if (curTotal > 0) {
      var pctCosto = ((curContable['Costo'] || 0) / curTotal) * 100;
      insights.push({ tipo: 'info', icono: 'layers',
        titulo: pctCosto.toFixed(0) + '% costos directos · ' + (100 - pctCosto).toFixed(0) + '% gastos',
        detalle: 'Los costos suben con más pacientes (variable sano); los gastos fijos no — ahí está el ahorro estructural.' });
    }
    // Candidatos a recortar: crecieron >25% y pesan >3%
    var candidatos = subArr.filter(function (s) { return s.crec != null && s.crec > 25 && s.pct > 3; }).slice(0, 3);
    if (candidatos.length) {
      insights.push({ tipo: 'warn', icono: 'scissors',
        titulo: 'Candidatos a revisar para ahorrar',
        detalle: candidatos.map(function (c) { return c.subtipo + ' (+' + c.crec.toFixed(0) + '%)'; }).join(' · ') + '. Crecieron fuerte y pesan en el total.' });
    }
    // Concentración de proveedor
    if (provArr.length && provArr[0].pct > 20) {
      insights.push({ tipo: 'info', icono: 'briefcase',
        titulo: provArr[0].proveedor + ' es ' + provArr[0].pct.toFixed(0) + '% de tu gasto',
        detalle: 'Alta dependencia de un proveedor — buen punto para negociar volumen o buscar alternativa.' });
    }

    return {
      ok: true,
      anioActual: anioActual, anioAnterior: anioAnterior,
      totalActual: curTotal, totalAnterior: prevTotal, crecTotal: crecTotal,
      porAnio: porAnio,
      contable: curContable,
      fijo: fijo, variable: variable,
      meses: mesesArr,
      subtipos: subArr,
      proveedores: provArr,
      insights: insights
    };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}
