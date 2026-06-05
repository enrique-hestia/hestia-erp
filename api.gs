var SHEET_ID = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4';

/* ══════════════════════════════════════════════════════════════
   doGet — Enrutador principal
   ?action=menu              → menú + periodos (carga inicial, ligero)
   ?action=view&view=X&periodo=Y → datos de esa vista específica
   ══════════════════════════════════════════════════════════════ */
function doGet(e) {
  try {
    var ss      = SpreadsheetApp.openById(SHEET_ID);
    var action  = (e && e.parameter.action)  || 'menu';
    var periodo = (e && e.parameter.periodo) || getDefaultPeriodo(ss);
    var view    = (e && e.parameter.view)    || 'resumen';

    if (action === 'menu') {
      return jsonResponse({
        menu:     readMenu(ss),
        periodos: readPeriodos(ss),
        periodo:  periodo
      });
    }

    if (action === 'insert') {
      return jsonResponse(insertRow(ss, e));
    }

    // action === 'view'
    return jsonResponse(readViewData(ss, view, periodo));

  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Periodo por defecto: el primero de la hoja Periodos ─── */
function getDefaultPeriodo(ss) {
  var periodos = readPeriodos(ss);
  return periodos.length ? periodos[0].id : '2026-Q2';
}

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
   fuente = 'mensual' → datos financieros filtrados por periodo
   fuente = NombreHoja → lee esa hoja como tabla de captura
   ══════════════════════════════════════════════════════════════ */
function readViewData(ss, viewId, periodo) {
  var menu   = readMenu(ss);
  var item   = null;
  for (var i = 0; i < menu.length; i++) {
    if (menu[i].id === viewId) { item = menu[i]; break; }
  }
  var fuente = item ? item.fuente : 'mensual';

  if (!fuente || fuente === 'mensual') {
    return readMensualData(ss, periodo, viewId);
  }

  // Captura dinámica: lee la hoja por nombre
  return readCapturaData(ss, fuente, periodo, viewId);
}

/* ── Datos financieros completos ─── */
function readMensualData(ss, periodo, viewId) {
  return {
    view:          viewId,
    periodo:       periodo,
    todos:         readMensual(ss, 'Mensual_Todos', periodo),
    local:         readMensual(ss, 'Mensual_Local', periodo),
    internacional: readMensual(ss, 'Mensual_Internacional', periodo),
    servicios:     readServicios(ss),
    funnel:        readFunnel(ss),
    alertas:       readAlertas(ss),
    donut:         readDonut(ss),
    cashflow:      readCashFlow(ss, periodo),
    costos:        readCostos(ss),
    paisesOrigen:  readPaisesOrigen(ss, periodo),
    updated:       new Date().toISOString()
  };
}

/* ── Datos de una hoja de captura genérica ─── */
function readCapturaData(ss, nombreHoja, periodo, viewId) {
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    return { view: viewId, periodo: periodo, headers: [], rows: [],
             error: 'Hoja "' + nombreHoja + '" no encontrada en el Spreadsheet.' };
  }
  var allRows = hoja.getDataRange().getValues();
  if (allRows.length < 1) return { view: viewId, periodo: periodo, headers: [], rows: [] };

  // Cabecera: row 0, columnas desde B (índice 1 en adelante)
  var headers = allRows[0].slice(1).map(function(h) {
    return String(h).trim();
  });

  // Filas filtradas por periodo (columna A)
  var rows = allRows.slice(1)
    .filter(function(r) { return String(r[0]).trim() === periodo; })
    .map(function(r) {
      var obj = { _periodo: String(r[0]) };
      headers.forEach(function(h, i) { obj[h] = r[i + 1]; });
      return obj;
    });

  return {
    view:    viewId,
    periodo: periodo,
    fuente:  nombreHoja,
    headers: headers,
    rows:    rows,
    updated: new Date().toISOString()
  };
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

function readMensual(ss, sheetName, periodo) {
  var rows = ss.getSheetByName(sheetName).getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) { return String(r[0]) === periodo; });
  return {
    meses:    data.map(function(r) { return String(r[1]); }),
    ingresos: data.map(function(r) { return Number(r[2]); }),
    gastos:   data.map(function(r) { return Number(r[3]); }),
    ciclos:   data.map(function(r) { return Number(r[4]); }),
    cac:      data.map(function(r) { return Number(r[5]); }),
    margen:   data.map(function(r) { return Number(r[6]); })
  };
}

function readCashFlow(ss, periodo) {
  var rows = ss.getSheetByName('CashFlow').getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) { return String(r[0]) === periodo; });
  return {
    meses: data.map(function(r) { return String(r[1]); }),
    flujo: data.map(function(r) { return Number(r[2]); })
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

function readPaisesOrigen(ss, periodo) {
  var hoja = ss.getSheetByName('PaisesOrigen');
  if (!hoja) return { labels: [], data: [], colors: [] };
  var rows = hoja.getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) { return String(r[0]) === periodo; });
  return {
    labels: data.map(function(r) { return String(r[1]); }),
    data:   data.map(function(r) { return Number(r[2]); }),
    colors: data.map(function(r) { return String(r[3]); })
  };
}

/* ══════════════════════════════════════════════════════════════
   INSERT ROW — agrega una fila al final de la hoja indicada
   Params: sheet, periodo, + una clave por columna (según cabecera)
   ══════════════════════════════════════════════════════════════ */
function insertRow(ss, e) {
  var sheetName = (e && e.parameter.sheet) || '';
  var hoja = ss.getSheetByName(sheetName);
  if (!hoja) return { error: 'Hoja "' + sheetName + '" no encontrada.' };

  var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var row = headers.map(function(h, i) {
    if (i === 0) return (e && e.parameter.periodo) || '';
    var key = String(h).trim();
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
