var SHEET_ID = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4';

/* ══════════════════════════════════════════════════════
   doGet — Punto de entrada de la Web App
   Parámetro opcional: ?periodo=2026-Q2
   ══════════════════════════════════════════════════════ */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var periodos = readPeriodos(ss);

    // Si no se pasa periodo, usar el primero de la lista (el más reciente)
    var periodo = (e && e.parameter && e.parameter.periodo)
                  ? e.parameter.periodo
                  : periodos[0].id;

    var result = {
      periodo:       periodo,
      periodos:      periodos,
      todos:         readMensual(ss, 'Mensual_Todos', periodo),
      local:         readMensual(ss, 'Mensual_Local', periodo),
      internacional: readMensual(ss, 'Mensual_Internacional', periodo),
      servicios:     readServicios(ss),
      funnel:        readFunnel(ss),
      alertas:       readAlertas(ss),
      donut:         readDonut(ss),
      cashflow:      readCashFlow(ss, periodo),
      costos:        readCostos(ss),
      inventarios:   readHojaCaptura(ss, 'Inventarios', periodo, ['producto','categoria','stock_actual','stock_minimo','costo_unitario','proveedor','ultima_compra']),
      laboratorios:  readHojaCaptura(ss, 'Laboratorios', periodo, ['fecha','estudio','paciente_id','tecnico','resultado','costo','estado']),
      updated:       new Date().toISOString()
    };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ── Lee la hoja Periodos ─── */
function readPeriodos(ss) {
  var rows = ss.getSheetByName('Periodos').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { id: String(r[0]), label: String(r[1]), orden: Number(r[2]) };
  });
}

/* ── Lee una hoja mensual filtrada por periodo ─── */
function readMensual(ss, sheetName, periodo) {
  var rows = ss.getSheetByName(sheetName).getDataRange().getValues();
  // Columna 0 = Periodo, 1 = Mes, 2 = Ingresos, 3 = Gastos, 4 = Ciclos, 5 = CAC, 6 = Margen
  var data = rows.slice(1).filter(function(r) {
    return String(r[0]) === periodo;
  });
  return {
    meses:    data.map(function(r) { return String(r[1]); }),
    ingresos: data.map(function(r) { return Number(r[2]); }),
    gastos:   data.map(function(r) { return Number(r[3]); }),
    ciclos:   data.map(function(r) { return Number(r[4]); }),
    cac:      data.map(function(r) { return Number(r[5]); }),
    margen:   data.map(function(r) { return Number(r[6]); }),
  };
}

/* ── Lee CashFlow filtrado por periodo ─── */
function readCashFlow(ss, periodo) {
  var rows = ss.getSheetByName('CashFlow').getDataRange().getValues();
  // Columna 0 = Periodo, 1 = Mes, 2 = Flujo
  var data = rows.slice(1).filter(function(r) {
    return String(r[0]) === periodo;
  });
  return {
    meses: data.map(function(r) { return String(r[1]); }),
    flujo: data.map(function(r) { return Number(r[2]); }),
  };
}

/* ── Las siguientes no filtran por periodo (datos fijos) ─── */
function readServicios(ss) {
  var rows = ss.getSheetByName('Servicios').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { name: String(r[0]), color: String(r[1]), ingresos: String(r[2]), margen: Number(r[3]), meta: Number(r[4]) };
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
    colors: data.map(function(r) { return String(r[2]); }),
  };
}
/* ── Lee cualquier hoja de Captura filtrada por periodo ─── */
function readHojaCaptura(ss, nombreHoja, periodo, campos) {
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return [];  // Si la hoja no existe aún, devuelve vacío
  var rows = hoja.getDataRange().getValues();
  // Columna 0 = Periodo, resto = campos
  return rows.slice(1)
    .filter(function(r) { return String(r[0]) === periodo; })
    .map(function(r) {
      var obj = {};
      campos.forEach(function(campo, i) { obj[campo] = r[i + 1]; });
      return obj;
    });
}

function readCostos(ss) {
  var rows = ss.getSheetByName('DistribucionCostos').getDataRange().getValues();
  var data = rows.slice(1);
  return {
    labels: data.map(function(r) { return String(r[0]); }),
    data:   data.map(function(r) { return Number(r[1]); }),
    colors: data.map(function(r) { return String(r[2]); }),
  };
}

/* ══════════════════════════════════════════════════════
   setupSheets — Ejecutar UNA sola vez para migrar
   estructura al nuevo formato con columna Periodo
   ══════════════════════════════════════════════════════ */
function setupSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Hoja: Periodos ──────────────────────────────────
  crearHoja(ss, 'Periodos', [
    ['ID',          'Label',       'Orden'],
    ['2026-Q2',     'Q2 2026',     1],
    ['2026-Q1',     'Q1 2026',     2],
    ['2025-Anual',  '2025 Anual',  3],
  ]);

  // ── Hoja: Mensual_Todos (nueva estructura con Periodo) ──
  crearHoja(ss, 'Mensual_Todos', [
    ['Periodo',   'Mes', 'Ingresos', 'Gastos', 'Ciclos', 'CAC', 'Margen'],
    // Q2 2026
    ['2026-Q2',   'Nov', 1800, 1150, 38, 54, 36],
    ['2026-Q2',   'Dic', 1950, 1200, 41, 53, 38],
    ['2026-Q2',   'Ene', 2100, 1280, 43, 52, 39],
    ['2026-Q2',   'Feb', 2250, 1350, 45, 51, 38],
    ['2026-Q2',   'Mar', 2180, 1330, 44, 52, 39],
    ['2026-Q2',   'Abr', 2400, 1490, 47, 51, 38],
    // Q1 2026
    ['2026-Q1',   'Jul', 1600, 1050, 34, 56, 34],
    ['2026-Q1',   'Ago', 1680, 1080, 35, 55, 36],
    ['2026-Q1',   'Sep', 1720, 1100, 36, 55, 36],
    ['2026-Q1',   'Oct', 1750, 1120, 37, 54, 36],
    ['2026-Q1',   'Nov', 1790, 1140, 37, 54, 36],
    ['2026-Q1',   'Dic', 1800, 1150, 38, 54, 36],
    // 2025 Anual
    ['2025-Anual','Ene', 1200, 800,  28, 60, 33],
    ['2025-Anual','Feb', 1250, 820,  29, 59, 34],
    ['2025-Anual','Mar', 1300, 850,  30, 58, 35],
    ['2025-Anual','Abr', 1350, 870,  31, 58, 36],
    ['2025-Anual','May', 1400, 890,  32, 57, 36],
    ['2025-Anual','Jun', 1450, 910,  33, 57, 37],
    ['2025-Anual','Jul', 1480, 930,  34, 56, 37],
    ['2025-Anual','Ago', 1510, 950,  35, 56, 37],
    ['2025-Anual','Sep', 1540, 970,  35, 55, 37],
    ['2025-Anual','Oct', 1570, 990,  36, 55, 37],
    ['2025-Anual','Nov', 1600,1010,  37, 55, 37],
    ['2025-Anual','Dic', 1630,1030,  37, 55, 37],
  ]);

  // ── Hoja: Mensual_Local ─────────────────────────────
  crearHoja(ss, 'Mensual_Local', [
    ['Periodo',   'Mes', 'Ingresos', 'Gastos', 'Ciclos', 'CAC', 'Margen'],
    ['2026-Q2',   'Nov', 1100, 730,  26, 42, 34],
    ['2026-Q2',   'Dic', 1200, 750,  28, 41, 37],
    ['2026-Q2',   'Ene', 1250, 800,  29, 40, 36],
    ['2026-Q2',   'Feb', 1380, 850,  31, 40, 38],
    ['2026-Q2',   'Mar', 1320, 820,  30, 41, 38],
    ['2026-Q2',   'Abr', 1470, 920,  32, 40, 37],
    ['2026-Q1',   'Jul',  980, 660,  22, 44, 33],
    ['2026-Q1',   'Ago', 1010, 680,  23, 44, 33],
    ['2026-Q1',   'Sep', 1040, 700,  24, 43, 33],
    ['2026-Q1',   'Oct', 1060, 710,  24, 43, 33],
    ['2026-Q1',   'Nov', 1080, 720,  25, 43, 33],
    ['2026-Q1',   'Dic', 1100, 730,  26, 42, 34],
    ['2025-Anual','Ene',  750, 510,  18, 48, 32],
    ['2025-Anual','Feb',  780, 525,  19, 47, 33],
    ['2025-Anual','Mar',  810, 540,  20, 47, 33],
    ['2025-Anual','Abr',  840, 555,  20, 46, 34],
    ['2025-Anual','May',  870, 570,  21, 46, 34],
    ['2025-Anual','Jun',  900, 585,  21, 46, 35],
    ['2025-Anual','Jul',  920, 600,  22, 45, 35],
    ['2025-Anual','Ago',  940, 615,  22, 45, 35],
    ['2025-Anual','Sep',  960, 628,  23, 45, 35],
    ['2025-Anual','Oct',  975, 640,  23, 45, 34],
    ['2025-Anual','Nov',  990, 652,  24, 44, 34],
    ['2025-Anual','Dic', 1000, 660,  24, 44, 34],
  ]);

  // ── Hoja: Mensual_Internacional ─────────────────────
  crearHoja(ss, 'Mensual_Internacional', [
    ['Periodo',   'Mes', 'Ingresos', 'Gastos', 'Ciclos', 'CAC', 'Margen'],
    ['2026-Q2',   'Nov',  700, 420, 12, 72, 40],
    ['2026-Q2',   'Dic',  750, 450, 13, 70, 40],
    ['2026-Q2',   'Ene',  850, 480, 14, 68, 43],
    ['2026-Q2',   'Feb',  870, 500, 14, 68, 43],
    ['2026-Q2',   'Mar',  860, 510, 14, 70, 41],
    ['2026-Q2',   'Abr',  930, 570, 15, 69, 39],
    ['2026-Q1',   'Jul',  620, 390, 12, 74, 37],
    ['2026-Q1',   'Ago',  650, 400, 12, 73, 38],
    ['2026-Q1',   'Sep',  670, 410, 13, 73, 39],
    ['2026-Q1',   'Oct',  680, 415, 13, 72, 39],
    ['2026-Q1',   'Nov',  700, 420, 12, 72, 40],
    ['2026-Q1',   'Dic',  700, 420, 12, 72, 40],
    ['2025-Anual','Ene',  450, 290, 10, 80, 36],
    ['2025-Anual','Feb',  470, 295, 10, 79, 37],
    ['2025-Anual','Mar',  490, 310, 10, 79, 37],
    ['2025-Anual','Abr',  510, 315, 11, 78, 38],
    ['2025-Anual','May',  530, 320, 11, 78, 40],
    ['2025-Anual','Jun',  550, 325, 12, 77, 41],
    ['2025-Anual','Jul',  560, 350, 12, 77, 38],
    ['2025-Anual','Ago',  570, 335, 13, 77, 41],
    ['2025-Anual','Sep',  580, 342, 13, 76, 41],
    ['2025-Anual','Oct',  595, 350, 13, 76, 41],
    ['2025-Anual','Nov',  610, 358, 13, 75, 41],
    ['2025-Anual','Dic',  630, 370, 13, 75, 41],
  ]);

  // ── Hoja: CashFlow (nueva estructura con Periodo) ───
  crearHoja(ss, 'CashFlow', [
    ['Periodo',   'Mes', 'Flujo_MXN_K'],
    ['2026-Q2',   'Nov', 420],
    ['2026-Q2',   'Dic', 580],
    ['2026-Q2',   'Ene', 650],
    ['2026-Q2',   'Feb', 720],
    ['2026-Q2',   'Mar', 680],
    ['2026-Q2',   'Abr', 740],
    ['2026-Q1',   'Jul', 310],
    ['2026-Q1',   'Ago', 340],
    ['2026-Q1',   'Sep', 370],
    ['2026-Q1',   'Oct', 390],
    ['2026-Q1',   'Nov', 410],
    ['2026-Q1',   'Dic', 420],
    ['2025-Anual','Ene', 180],
    ['2025-Anual','Feb', 200],
    ['2025-Anual','Mar', 220],
    ['2025-Anual','Abr', 245],
    ['2025-Anual','May', 260],
    ['2025-Anual','Jun', 280],
    ['2025-Anual','Jul', 290],
    ['2025-Anual','Ago', 300],
    ['2025-Anual','Sep', 310],
    ['2025-Anual','Oct', 320],
    ['2025-Anual','Nov', 335],
    ['2025-Anual','Dic', 350],
  ]);

  Logger.log('✅ Migración completada con columna Periodo');
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
