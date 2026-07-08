/* ==============================================================
   medical.gs — Módulo Médico
   --------------------------------------------------------------
   Quirófano, Reporte Ejecutivo, Dashboard Medicamentos
   Proyecto Google Apps Script — Hestia Fertility ERP
   Todas las constantes vienen de config.gs (mismo proyecto)
   ============================================================== */

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