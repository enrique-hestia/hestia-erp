var SHEET_ID      = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4';
var API_VERSION   = 'v2026-06-08-G';  // Actualizar al redesplegar para verificar versión

// Mapeo: nombre de pestaña → ID del spreadsheet externo donde se lee/escribe
// Agregar aquí cualquier hoja de captura futura
var CAPTURA_SHEETS = {
  'Medicamentos': '1fiuUtw-sg2ELNxq9bCjaOtRz1n87wuVi8IOQYzEi8tM',
  'Insumos':      '1hYmIl4gSTVrvghP7KY0y0dC200o8w0zShXj63zP-TrQ',
  'Pacientes':    '1uoQU-vbefxWwaLxJyTFT25gj7Nr2223WISa3tqH-Rio'
};
// Fallback si la hoja no está en el mapeo (usa el sheet principal de Hestia ERP)
var CAPTURA_SHEET_ID_DEFAULT = SHEET_ID;

/* ══════════════════════════════════════════════════════════════
   doGet — Enrutador principal
   ?action=menu                              → menú (carga inicial)
   ?action=view&view=X&fechaInicio=Y&fechaFin=Z → datos de la vista
   ?action=insert&sheet=X&...campos          → inserta fila
   ══════════════════════════════════════════════════════════════ */
function doGet(e) {
  try {
    var ss     = SpreadsheetApp.openById(SHEET_ID);
    var action = (e && e.parameter.action) || 'menu';
    var view   = (e && e.parameter.view)   || 'resumen';

    // Rango de fechas — default: últimos 6 meses
    var hoy        = new Date();
    var defInicio  = new Date(hoy.getFullYear(), hoy.getMonth() - 5, 1);
    var defFin     = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);
    var fechaInicio = (e && e.parameter.fechaInicio) || fmtDate(defInicio);
    var fechaFin    = (e && e.parameter.fechaFin)    || fmtDate(defFin);

    if (action === 'menu') {
      return jsonResponse({
        menu:        readMenu(ss),
        fechaInicio: fechaInicio,
        fechaFin:    fechaFin,
        version:     API_VERSION
      });
    }

    if (action === 'insert') {
      return jsonResponse(insertRow(ss, e));
    }

    // Debug: ?action=debug&sheet=NombreHoja
    if (action === 'debug') {
      var sheetName = (e && e.parameter.sheet) || '';
      var capturaId = CAPTURA_SHEETS[sheetName] || CAPTURA_SHEET_ID_DEFAULT;
      try {
        var ssDeb = SpreadsheetApp.openById(capturaId);
        var tabs  = ssDeb.getSheets().map(function(s){ return s.getName(); });
        return jsonResponse({ sheetName: sheetName, spreadsheetId: capturaId,
                              tabsFound: tabs, capturaSheets: CAPTURA_SHEETS });
      } catch(ex) {
        return jsonResponse({ error: ex.message, sheetName: sheetName, spreadsheetId: capturaId });
      }
    }

    // options: ?action=options&sheet=NombreHoja → lee validaciones dropdown de la hoja
    if (action === 'options') {
      var sheetName = (e && e.parameter.sheet) || 'Pacientes';
      var capturaId = CAPTURA_SHEETS[sheetName] || CAPTURA_SHEET_ID_DEFAULT;
      try {
        var ssOpt  = SpreadsheetApp.openById(capturaId);
        var shOpt  = ssOpt.getSheetByName(sheetName);
        if (!shOpt) return jsonResponse({ error: 'Hoja no encontrada: ' + sheetName });
        var lastCol  = shOpt.getLastColumn();
        var headers  = shOpt.getRange(1, 1, 1, lastCol).getValues()[0];
        var targets  = ['Origen', 'Canal', 'Médico Tratante', 'País'];
        var options  = {};
        // Leer opciones desde pestaña "Opciones" (una columna por campo dropdown)
        var shOpciones = ssOpt.getSheetByName('Opciones');
        if (shOpciones) {
          var optHeaders = shOpciones.getRange(1, 1, 1, shOpciones.getLastColumn()).getValues()[0];
          var optData    = shOpciones.getRange(2, 1, Math.max(shOpciones.getLastRow() - 1, 1), shOpciones.getLastColumn()).getValues();
          targets.forEach(function(colName) {
            var colIdx = -1;
            for (var i = 0; i < optHeaders.length; i++) {
              if (String(optHeaders[i]).trim() === colName) { colIdx = i; break; }
            }
            if (colIdx === -1) { options[colName] = []; return; }
            options[colName] = optData
              .map(function(row) { return String(row[colIdx]).trim(); })
              .filter(function(v) { return v && v !== '' && v !== 'undefined'; });
          });
        }
        return jsonResponse({ options: options });
      } catch(ex) {
        return jsonResponse({ error: ex.message });
      }
    }

    // nextid: ?action=nextid&sheet=Pacientes&prefix=HEC → siguiente ID disponible
    if (action === 'nextid') {
      var sheetName = (e && e.parameter.sheet)  || 'Pacientes';
      var prefix    = (e && e.parameter.prefix) || 'HEC';
      var capturaId = CAPTURA_SHEETS[sheetName] || CAPTURA_SHEET_ID_DEFAULT;
      try {
        var ssNid = SpreadsheetApp.openById(capturaId);
        var shNid = ssNid.getSheetByName(sheetName);
        if (!shNid || shNid.getLastRow() < 2) {
          return jsonResponse({ nextId: prefix + '-001' });
        }
        var ids = shNid.getRange(2, 1, shNid.getLastRow() - 1, 1).getValues()
          .map(function(r) { return String(r[0]); });
        var maxNum = 0;
        ids.forEach(function(id) {
          var m = id.match(/(\d+)$/);
          if (m) maxNum = Math.max(maxNum, parseInt(m[1]));
        });
        return jsonResponse({ nextId: prefix + '-' + String(maxNum + 1).padStart(3, '0') });
      } catch(ex) {
        return jsonResponse({ nextId: prefix + '-001', error: ex.message });
      }
    }

    // update: ?action=update&sheet=X&rowNum=N&Campo=valor → actualiza fila en Sheets
    if (action === 'update') {
      var sheetName = (e && e.parameter.sheet)  || '';
      var rowNum    = parseInt((e && e.parameter.rowNum) || '0');
      if (!sheetName || !rowNum) return jsonResponse({ error: 'sheet y rowNum son requeridos' });
      var capturaId = CAPTURA_SHEETS[sheetName] || CAPTURA_SHEET_ID_DEFAULT;
      try {
        var ssUpd = SpreadsheetApp.openById(capturaId);
        var shUpd = ssUpd.getSheetByName(sheetName);
        if (!shUpd) return jsonResponse({ error: 'Hoja no encontrada: ' + sheetName });
        var hdrs = shUpd.getRange(1, 1, 1, shUpd.getLastColumn()).getValues()[0];
        var cur  = shUpd.getRange(rowNum, 1, 1, hdrs.length).getValues()[0];
        var newRow = hdrs.map(function(h, i) {
          var key = String(h).trim();
          return (e.parameter[key] !== undefined) ? e.parameter[key] : cur[i];
        });
        shUpd.getRange(rowNum, 1, 1, hdrs.length).setValues([newRow]);
        return jsonResponse({ success: true, rowNum: rowNum });
      } catch(ex) {
        return jsonResponse({ error: ex.message });
      }
    }

    // action === 'view'
    return jsonResponse(readViewData(ss, view, fechaInicio, fechaFin));

  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

/* Formatea Date como YYYY-MM-DD */
function fmtDate(d) {
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  return d.getFullYear() + '-' + mm + '-' + dd;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── getDefaultPeriodo mantenido por compatibilidad (ya no se usa) ─── */
function getDefaultPeriodo(ss) { return ''; }

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
function readViewData(ss, viewId, fechaInicio, fechaFin) {
  var menu = readMenu(ss);
  var item = null;
  for (var i = 0; i < menu.length; i++) {
    if (menu[i].id === viewId) { item = menu[i]; break; }
  }
  var fuente = item ? item.fuente : 'mensual';

  if (!fuente || fuente === 'mensual') {
    return readMensualData(ss, fechaInicio, fechaFin, viewId);
  }

  return readCapturaData(ss, fuente, viewId, fechaInicio, fechaFin);
}

/* ── Datos financieros completos (filtrados por rango de fechas) ─── */
function readMensualData(ss, fechaInicio, fechaFin, viewId) {
  var label = fechaInicio.slice(0,7) + ' → ' + fechaFin.slice(0,7);
  return {
    view:          viewId,
    periodo:       label,          // string descriptivo para subtítulos
    fechaInicio:   fechaInicio,
    fechaFin:      fechaFin,
    todos:         readMensual(ss, 'Mensual_Todos',         fechaInicio, fechaFin),
    local:         readMensual(ss, 'Mensual_Local',         fechaInicio, fechaFin),
    internacional: readMensual(ss, 'Mensual_Internacional', fechaInicio, fechaFin),
    servicios:     readServicios(ss),
    funnel:        readFunnel(ss),
    alertas:       readAlertas(ss),
    donut:         readDonut(ss),
    cashflow:      readCashFlow(ss, fechaInicio, fechaFin),
    costos:        readCostos(ss),
    paisesOrigen:  readPaisesOrigen(ss, fechaInicio, fechaFin),
    updated:       new Date().toISOString()
  };
}

/* ── Datos de hoja de captura filtrados por rango de fechas ─────────
   Espera col A = Periodo, col B = Fecha (YYYY-MM-DD), resto = datos.
   Si fechaInicio/fechaFin están vacíos devuelve todas las filas.
   ──────────────────────────────────────────────────────────────── */
function readCapturaData(ss, nombreHoja, viewId, fechaInicio, fechaFin) {
  var capturaId = CAPTURA_SHEETS[nombreHoja] || CAPTURA_SHEET_ID_DEFAULT;
  var ssCap = SpreadsheetApp.openById(capturaId);
  var hoja  = ssCap.getSheetByName(nombreHoja);
  if (!hoja) {
    return { view: viewId, headers: [], rows: [],
             error: 'Hoja "' + nombreHoja + '" no encontrada.' };
  }
  var allRows = hoja.getDataRange().getValues();
  if (allRows.length < 1) return { view: viewId, headers: [], rows: [] };

  var headerRow = allRows[0];
  // Detectar si la hoja tiene estructura Periodo en col A
  var tienePeriodo = String(headerRow[0]).trim().toLowerCase() === 'periodo';
  // Detectar si además tiene Fecha en col B (solo si tiene Periodo)
  var tieneFecha   = tienePeriodo && String(headerRow[1]).trim().toLowerCase() === 'fecha';
  // colStart: desde dónde empiezan las columnas visibles en la tabla
  // - Con Periodo+Fecha: desde col C (índice 2)
  // - Con solo Periodo:  desde col B (índice 1)
  // - Sin Periodo:       desde col A (índice 0) — incluye todo
  var colStart = tieneFecha ? 2 : (tienePeriodo ? 1 : 0);
  var headers = headerRow.slice(colStart).map(function(h) { return String(h).trim(); });

  var dataRows = allRows.slice(1);

  // Filtrar por fecha si la hoja tiene columna Fecha y se pasó rango
  if (tieneFecha && fechaInicio && fechaFin) {
    dataRows = dataRows.filter(function(r) {
      var f = String(r[1]).trim(); // col B = Fecha
      return f >= fechaInicio && f <= fechaFin;
    });
    // Ordenar por fecha ascendente
    dataRows.sort(function(a, b) { return String(a[1]) < String(b[1]) ? -1 : 1; });
  }

  // Guardar número de fila original antes de filtrar (para edición posterior)
  var dataRowsWithNum = allRows.slice(1).map(function(r, i) {
    return { data: r, rowNum: i + 2 };
  });
  if (tieneFecha && fechaInicio && fechaFin) {
    dataRowsWithNum = dataRowsWithNum.filter(function(item) {
      var f = String(item.data[1]).trim();
      return f >= fechaInicio && f <= fechaFin;
    });
    dataRowsWithNum.sort(function(a, b) { return String(a.data[1]) < String(b.data[1]) ? -1 : 1; });
  } else if (!tieneFecha) {
    dataRowsWithNum = dataRowsWithNum.filter(function(item) {
      return item.data.some(function(c) { return String(c).trim() !== ''; });
    });
  }

  var rows = dataRowsWithNum.map(function(item) {
    var r = item.data;
    var obj = { _rowNum: item.rowNum, _periodo: String(r[0]), _fecha: tieneFecha ? String(r[1]) : '' };
    headers.forEach(function(h, i) {
      obj[h] = r[colStart + i];
      obj[h.toLowerCase()] = r[colStart + i];
    });
    return obj;
  });

  return { view: viewId, fuente: nombreHoja, headers: headers, rows: rows,
           updated: new Date().toISOString() };
}

/* ══════════════════════════════════════════════════════════════
   LECTORES INDIVIDUALES
   ══════════════════════════════════════════════════════════════ */
function readPeriodos(ss) {
  var rows = ss.getSheetByName('Periodos').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { id: String(r[0]), label: String(r[1]), orden: Number(r[2]) };
  });
}

/* Columnas esperadas (con la nueva col Fecha en B):
   A=Periodo | B=Fecha(YYYY-MM-DD) | C=Mes | D=Ingresos | E=Gastos | F=Ciclos | G=CAC | H=Margen */
function readMensual(ss, sheetName, fechaInicio, fechaFin) {
  var rows = ss.getSheetByName(sheetName).getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[1]).trim(); // col B = Fecha
    return f >= fechaInicio && f <= fechaFin;
  });
  // Ordenar por fecha ascendente
  data.sort(function(a, b) { return String(a[1]) < String(b[1]) ? -1 : 1; });
  return {
    meses:    data.map(function(r) { return String(r[2]); }),  // col C
    ingresos: data.map(function(r) { return Number(r[3]); }),  // col D
    gastos:   data.map(function(r) { return Number(r[4]); }),  // col E
    ciclos:   data.map(function(r) { return Number(r[5]); }),  // col F
    cac:      data.map(function(r) { return Number(r[6]); }),  // col G
    margen:   data.map(function(r) { return Number(r[7]); })   // col H
  };
}

/* Columnas CashFlow: A=Periodo | B=Fecha | C=Mes | D=Flujo_MXN */
function readCashFlow(ss, fechaInicio, fechaFin) {
  var rows = ss.getSheetByName('CashFlow').getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[1]).trim();
    return f >= fechaInicio && f <= fechaFin;
  });
  data.sort(function(a, b) { return String(a[1]) < String(b[1]) ? -1 : 1; });
  return {
    meses: data.map(function(r) { return String(r[2]); }),
    flujo: data.map(function(r) { return Number(r[3]); })
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

/* PaisesOrigen: A=Periodo | B=Fecha | C=Pais | D=Porcentaje | E=Color */
function readPaisesOrigen(ss, fechaInicio, fechaFin) {
  var hoja = ss.getSheetByName('PaisesOrigen');
  if (!hoja) return { labels: [], data: [], colors: [] };
  var rows = hoja.getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[1]).trim();
    return f >= fechaInicio && f <= fechaFin;
  });
  return {
    labels: data.map(function(r) { return String(r[2]); }),
    data:   data.map(function(r) { return Number(r[3]); }),
    colors: data.map(function(r) { return String(r[4]); })
  };
}

/* ══════════════════════════════════════════════════════════════
   INSERT ROW — agrega una fila al final de la hoja indicada
   Params: sheet, periodo, + una clave por columna (según cabecera)
   ══════════════════════════════════════════════════════════════ */
function insertRow(ss, e) {
  var sheetName = (e && e.parameter.sheet) || '';
  // Abrir el spreadsheet correcto según el mapeo
  var capturaId = CAPTURA_SHEETS[sheetName] || CAPTURA_SHEET_ID_DEFAULT;
  var ssCap = SpreadsheetApp.openById(capturaId);
  var hoja = ssCap.getSheetByName(sheetName);
  if (!hoja) return { error: 'Hoja "' + sheetName + '" no encontrada.' };

  var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var row = headers.map(function(h, i) {
    var key = String(h).trim();
    // Si col A se llama 'Periodo', llenarlo con el param periodo automáticamente
    if (i === 0 && key === 'Periodo') return (e && e.parameter.periodo) || '';
    return (e && e.parameter[key] !== undefined) ? e.parameter[key] : '';
  });

  hoja.appendRow(row);
  return { success: true };
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
   setupSheets — Ejecutar para crear/migrar todas las hojas
   ══════════════════════════════════════════════════════════════ */
function setupSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Hoja: Menu ──────────────────────────────────────────────
  // Columnas: ID | Padre | Label | Seccion | Icono | Orden | Tipo | Fuente | Activo
  crearHoja(ss, 'Menu', [
    ['ID',            'Padre',     'Label',            'Seccion',  'Icono',            'Orden', 'Tipo',  'Fuente',   'Activo'],
    // ── PANEL ──
    ['resumen',       '',          'Resumen General',  'PANEL',    'layout-dashboard', 1,       'vista', 'mensual',  true],
    ['finanzas',      '',          'Finanzas',         'PANEL',    'landmark',         2,       'grupo', '',         true],
    ['ingresos',      'finanzas',  'Ingresos',         '',         'trending-up',      1,       'vista', 'mensual',  true],
    ['gastos',        'finanzas',  'Gastos Operativos','',         'receipt',          2,       'vista', 'mensual',  true],
    ['costos',        'finanzas',  'Costos',           '',         'calculator',       3,       'vista', 'mensual',  true],
    ['pacientes',     '',          'Pacientes',        'PANEL',    'users',            3,       'vista', 'mensual',  true],
    // ── ANÁLISIS ──
    ['analisis',      '',          'Análisis',         'ANÁLISIS', 'bar-chart-2',      4,       'grupo', '',         true],
    ['servicios',     'analisis',  'Servicios',        '',         'flask-conical',    1,       'vista', 'mensual',  true],
    ['turismo',       'analisis',  'Turismo Médico',   '',         'plane',            2,       'vista', 'mensual',  true],
    ['rentabilidad',  'analisis',  'Rentabilidad',     '',         'percent',          3,       'vista', 'mensual',  true],
    // ── CAPTURA ──
    ['captura',       '',          'Captura',          'CAPTURA',  'database',         5,       'grupo', '',         true],
    ['inventarios',   'captura',   'Inventarios',      '',         'package',          1,       'vista', 'Inventarios',  true],
    ['laboratorios',  'captura',   'Laboratorios',     '',         'microscope',       2,       'vista', 'Laboratorios', true],
    // ── CONFIG ──
    ['ajustes',       '',          'Ajustes',          'CONFIG',   'settings',         6,       'vista', '',         true],
  ]);

  // ── Hoja: Periodos ──────────────────────────────────────────
  crearHoja(ss, 'Periodos', [
    ['ID',         'Label',      'Orden'],
    ['2026-Q2',    'Q2 2026',    1],
    ['2026-Q1',    'Q1 2026',    2],
    ['2025-Anual', '2025 Anual', 3],
  ]);

  Logger.log('✅ setupSheets completado');
}

function crearHoja(ss, nombre, datos) {
  var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  h.clearContents();
  h.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
  var header = h.getRange(1, 1, 1, datos[0].length);
  header.setFontWeight('bold');
  header.setBackground('#fce8f0');
  h.setFrozenRows(1);
  h.autoResizeColumns(1, datos[0].length);
}
