/* ==============================================================
   capture.gs — Módulo Captura & Datos Mensuales
   --------------------------------------------------------------
   Insert/Update genérico, Caja Chica, Mensual, Captura
   Proyecto Google Apps Script — Hestia Fertility ERP
   Todas las constantes vienen de config.gs (mismo proyecto)
   ============================================================== */

function readMensualData(ss, fechaInicio, fechaFin, viewId, sucursal) {
  sucursal = sucursal || 'Todas';
  var label = fechaInicio.slice(0,7) + ' → ' + fechaFin.slice(0,7);
  return {
    view:          viewId,
    periodo:       label,
    fechaInicio:   fechaInicio,
    fechaFin:      fechaFin,
    todos:         readMensual(ss, 'Mensual_Todos',         fechaInicio, fechaFin, sucursal),
    local:         readMensual(ss, 'Mensual_Local',         fechaInicio, fechaFin, sucursal),
    internacional: readMensual(ss, 'Mensual_Internacional', fechaInicio, fechaFin, sucursal),
    servicios:     readServicios(ss),
    funnel:        readFunnel(ss),
    alertas:       readAlertas(ss),
    donut:         readDonut(ss),
    cashflow:      readCashFlow(ss, fechaInicio, fechaFin, sucursal),
    costos:        readCostos(ss),
    paisesOrigen:  readPaisesOrigen(ss, fechaInicio, fechaFin, sucursal),
    updated:       new Date().toISOString()
  };
}

/* ── Datos de hoja de captura — devuelve TODAS las filas (sin filtro de fechas)
   Las hojas de captura (Pacientes, Productos, Ent. Med, Estimulacion, etc.)
   muestran su contenido completo; el filtro de fechas aplica solo a datos financieros.
   ──────────────────────────────────────────────────────────────── */
function findSheet(ssCap, nombreHoja) {
  // 1. Nombre exacto
  var h = ssCap.getSheetByName(nombreHoja);
  if (h) return h;
  // 2. Búsqueda insensible a tildes y mayúsculas
  var normalize = function(s) {
    return s.trim().toLowerCase()
      .replace(/[áàäã]/g,'a').replace(/[éèë]/g,'e')
      .replace(/[íìï]/g,'i').replace(/[óòö]/g,'o')
      .replace(/[úùü]/g,'u').replace(/ñ/g,'n');
  };
  var target = normalize(nombreHoja);
  var sheets = ssCap.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalize(sheets[i].getName()) === target) return sheets[i];
  }
  return null;
}

/* Detecta la fila de encabezado de una hoja de captura (igual lógica que readCapturaData)
   y devuelve { headers, dataStart } usando posiciones absolutas de columna (incluye Periodo si existe). */
function getSheetHeaders(sheet) {
  var allRows = sheet.getDataRange().getValues();
  function countFilled(row) {
    return row.filter(function(c) { return String(c).trim() !== ''; }).length;
  }
  var r0 = allRows.length > 0 ? countFilled(allRows[0]) : 0;
  var r1 = allRows.length > 1 ? countFilled(allRows[1]) : 0;
  var headerRow, dataStart;

  if (r0 === 0) {
    headerRow = allRows[1] || [];
    dataStart = 2;
  } else if (r0 > 0 && r1 > 0) {
    var complementario = allRows[0].every(function(v, i) {
      var v0 = String(v).trim();
      var v1 = String((allRows[1][i] !== undefined ? allRows[1][i] : '')).trim();
      return !(v0 && v1);
    });
    if (complementario) {
      headerRow = allRows[0].map(function(v, i) {
        return String(v).trim() || String(allRows[1][i] !== undefined ? allRows[1][i] : '').trim();
      });
      dataStart = 2;
    } else {
      headerRow = allRows[0];
      dataStart = 1;
    }
  } else {
    headerRow = allRows[0] || [];
    dataStart = 1;
  }

  var headers = headerRow.map(function(h) { return String(h).trim(); });
  return { headers: headers, dataStart: dataStart };
}

/* insert: ?action=insert&sheet=X&Campo1=valor1&... → agrega una fila nueva al final */
function insertRow(ss, e) {
  var sheetName = (e && e.parameter.sheet) || '';
  if (SHEET_ALIASES[sheetName]) sheetName = SHEET_ALIASES[sheetName];
  if (!sheetName) return { error: 'sheet es requerido' };
  var capturaId = getCapturaId(sheetName);
  try {
    var ssIns = SpreadsheetApp.openById(capturaId);
    var shIns = findSheet(ssIns, sheetName);
    if (!shIns) return { error: 'Hoja no encontrada: ' + sheetName };
    var hdrInfo = getSheetHeaders(shIns);
    var hdrs = hdrInfo.headers;

    // Validación de duplicados: en Pacientes no se permite repetir el mismo nombre
    if (sheetName.trim().toLowerCase() === 'pacientes') {
      var nombreIdx = -1;
      for (var hi = 0; hi < hdrs.length; hi++) {
        if (hdrs[hi].toLowerCase().indexOf('nombre') > -1) { nombreIdx = hi; break; }
      }
      if (nombreIdx > -1) {
        var nombreNuevo = String(e.parameter[hdrs[nombreIdx]] || '').trim().toLowerCase();
        if (nombreNuevo) {
          var allData = shIns.getDataRange().getValues();
          for (var ri = hdrInfo.dataStart; ri < allData.length; ri++) {
            var existente = String(allData[ri][nombreIdx] || '').trim().toLowerCase();
            if (existente && existente === nombreNuevo) {
              return { error: 'Ya existe un paciente registrado con el nombre "' + e.parameter[hdrs[nombreIdx]] + '".', duplicado: true };
            }
          }
        }
      }
    }

    var newRow = hdrs.map(function(h) {
      return (h && e.parameter[h] !== undefined) ? e.parameter[h] : '';
    });
    shIns.appendRow(newRow);
    var rowNum = shIns.getLastRow();
    invalidateViewCache(sheetName);
    return { success: true, rowNum: rowNum };
  } catch(ex) {
    return { error: ex.message };
  }
}

/* ══════════════════════════════════════════════════════════════
   CAJA CHICA — hoja independiente con saldo corrido (columna TOTAL
   es fórmula en Sheets: Total_n = Total_(n-1) - Salida_n + Entrada_n).
   La pestaña activa lleva como nombre el año en curso (ej. "2026");
   muchas filas debajo de la última captura ya tienen la fórmula de
   TOTAL copiada hacia abajo en blanco, esperando captura futura —
   por eso nunca usamos appendRow() aquí, sino que buscamos la
   primera fila con FECHA y CONCEPTO vacíos para no romper el orden.
   ══════════════════════════════════════════════════════════════ */
function getCajaChicaSheet() {
  var ss = SpreadsheetApp.openById(CAJA_CHICA_SS_ID);
  var sh = ss.getSheetByName('Caja Chica');
  if (sh) return sh;
  var anioActual = String(new Date().getFullYear());
  sh = ss.getSheetByName(anioActual);
  if (sh) return sh;
  return ss.getSheets()[0]; // pestaña más reciente como último recurso
}

function readCajaChicaData() {
  try {
    var sh = getCajaChicaSheet();
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { movimientos: [], saldoInicial: 0, saldoFinal: 0 };
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iSalida   = headers.indexOf('SALIDA');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    function fmtFechaCC(v) {
      if (!v) return '';
      if (v instanceof Date) {
        var dd = String(v.getDate()).padStart(2,'0');
        var mm = String(v.getMonth()+1).padStart(2,'0');
        return dd + '/' + mm + '/' + v.getFullYear();
      }
      return String(v).trim();
    }
    var rows = [];
    var saldoInicial = 0;
    for (var r = 1; r < data.length; r++) {
      var row      = data[r];
      var fecha    = fmtFechaCC(row[iFecha]);
      var concepto = String(row[iConcepto] || '').trim();
      if (!fecha && !concepto) continue; // fila reservada (solo fórmula de TOTAL), aún sin capturar
      var salida   = Number(row[iSalida])  || 0;
      var entrada  = Number(row[iEntrada]) || 0;
      var total    = Number(row[iTotal])   || 0;
      var esRemanente = /^REMANENTE/i.test(fecha);
      if (esRemanente) saldoInicial = total;
      rows.push({
        _rowNum:    r + 1,
        fecha:      esRemanente ? '' : fecha,
        concepto:   esRemanente ? fecha : concepto,
        esRemanente: esRemanente,
        salida:     salida,
        entrada:    entrada,
        total:      total
      });
    }

    var saldoFinal = rows.length ? rows[rows.length - 1].total : saldoInicial;

    // Resumen de gasto por periodo (admite DD/MM/YYYY y MM/DD/YYYY mezclados en la hoja)
    function parseFechaMx(f) {
      var m = f.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (!m) return null;
      var a = parseInt(m[1], 10), b = parseInt(m[2], 10), y = parseInt(m[3], 10);
      var day = a, month = b;
      if (a > 12) { day = a; month = b; }
      else if (b > 12) { day = b; month = a; }
      return new Date(y, month - 1, day);
    }
    var hoy        = new Date();
    var inicioHoy  = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    var inicio7d   = new Date(inicioHoy.getTime() - 6 * 24 * 60 * 60 * 1000);
    var inicioMes  = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    var gastoHoy = 0, gastoSemana = 0, gastoMes = 0, ingresoMes = 0;
    rows.forEach(function(m) {
      if (m.esRemanente) return;
      var d = parseFechaMx(m.fecha);
      if (!d) return;
      if (d >= inicioMes) { gastoMes += m.salida; ingresoMes += m.entrada; }
      if (d >= inicio7d)  gastoSemana += m.salida;
      if (d >= inicioHoy) gastoHoy += m.salida;
    });

    return {
      saldoInicial: saldoInicial,
      saldoFinal:   saldoFinal,
      gastoHoy:     gastoHoy,
      gastoSemana:  gastoSemana,
      gastoMes:     gastoMes,
      ingresoMes:   ingresoMes,
      movimientos:  rows.slice().reverse(), // más reciente primero
      updated:      new Date().toISOString()
    };
  } catch(ex) {
    return { error: ex.message, movimientos: [], saldoInicial: 0, saldoFinal: 0 };
  }
}

function insertCajaChicaRow(e) {
  try {
    var sh = getCajaChicaSheet();
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iSalida   = headers.indexOf('SALIDA');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    var concepto = String(e.parameter['CONCEPTO'] || '').trim();
    var fecha    = String(e.parameter['FECHA']    || '').trim();
    var salida   = parseFloat(e.parameter['SALIDA'])  || 0;
    var entrada  = parseFloat(e.parameter['ENTRADA']) || 0;

    if (!concepto)            return { error: 'El concepto es requerido.' };
    if (!fecha)               return { error: 'La fecha es requerida.' };
    if (!salida && !entrada)  return { error: 'Captura un monto de salida o entrada.' };

    // Primera fila reservada (placeholder con fórmula de TOTAL) con FECHA y CONCEPTO vacíos
    var targetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (!String(data[r][iFecha] || '').trim() && !String(data[r][iConcepto] || '').trim()) {
        targetRow = r + 1; // fila 1-indexada en la hoja
        break;
      }
    }
    if (targetRow === -1) targetRow = sh.getLastRow() + 1;

    sh.getRange(targetRow, iFecha + 1).setValue(fecha);
    sh.getRange(targetRow, iConcepto + 1).setValue(concepto);
    if (salida)  sh.getRange(targetRow, iSalida + 1).setValue(salida);
    if (entrada) sh.getRange(targetRow, iEntrada + 1).setValue(entrada);
    SpreadsheetApp.flush();

    var nuevoTotal = sh.getRange(targetRow, iTotal + 1).getValue();
    return { success: true, rowNum: targetRow, saldoFinal: Number(nuevoTotal) || 0 };
  } catch(ex) {
    return { error: ex.message };
  }
}

function updateCajaChicaRow(body) {
  try {
    var sh      = getCajaChicaSheet();
    var rowNum  = parseInt(body.rowNum);
    if (!rowNum || rowNum < 2) return { error: 'Número de fila inválido.' };
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iSalida   = headers.indexOf('SALIDA');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    if (body.fecha)    sh.getRange(rowNum, iFecha + 1).setValue(body.fecha);
    if (body.concepto) sh.getRange(rowNum, iConcepto + 1).setValue(body.concepto);
    sh.getRange(rowNum, iSalida + 1).setValue(parseFloat(body.salida) || 0);
    sh.getRange(rowNum, iEntrada + 1).setValue(parseFloat(body.entrada) || 0);
    SpreadsheetApp.flush();
    var nuevoTotal = sh.getRange(rowNum, iTotal + 1).getValue();
    return { ok: true, rowNum: rowNum, saldoFinal: Number(nuevoTotal) || 0 };
  } catch(ex) {
    return { error: ex.message };
  }
}

function getCapturaId(nombreHoja) {
  if (CAPTURA_SHEETS[nombreHoja]) return CAPTURA_SHEETS[nombreHoja];
  // Búsqueda tolerante a tildes
  var normalize = function(s) {
    return s.trim().toLowerCase()
      .replace(/[áàäã]/g,'a').replace(/[éèë]/g,'e')
      .replace(/[íìï]/g,'i').replace(/[óòö]/g,'o')
      .replace(/[úùü]/g,'u').replace(/ñ/g,'n');
  };
  var target = normalize(nombreHoja);
  var keys = Object.keys(CAPTURA_SHEETS);
  for (var i = 0; i < keys.length; i++) {
    if (normalize(keys[i]) === target) return CAPTURA_SHEETS[keys[i]];
  }
  return CAPTURA_SHEET_ID_DEFAULT;
}

function readCapturaData(ss, nombreHoja, viewId, fechaInicio, fechaFin, sucursal) {
  sucursal = sucursal || 'Todas';
  var capturaId = getCapturaId(nombreHoja);
  var ssCap = SpreadsheetApp.openById(capturaId);
  var hoja  = findSheet(ssCap, nombreHoja);
  if (!hoja) {
    return { view: viewId, headers: [], rows: [],
             error: 'Hoja "' + nombreHoja + '" no encontrada.' };
  }
  var allRows = hoja.getDataRange().getValues();
  if (allRows.length < 1) return { view: viewId, headers: [], rows: [] };

  // ── Detección inteligente de fila de encabezado ────────────────
  // Algunas hojas tienen fila 1 vacía (FET, Insumos) o encabezados
  // divididos entre fila 1 y fila 2 (ART Lab). Se detecta automáticamente.
  function countFilled(row) {
    return row.filter(function(c) { return String(c).trim() !== ''; }).length;
  }
  var r0 = countFilled(allRows[0]);
  var r1 = allRows.length > 1 ? countFilled(allRows[1]) : 0;
  var headerRow, dataStart;

  if (r0 === 0) {
    // Fila 1 vacía → encabezado en fila 2 (FET, Insumos)
    headerRow  = allRows[1] || [];
    dataStart  = 2;
  } else if (r0 > 0 && r1 > 0) {
    // Ambas filas tienen datos → verificar si son encabezados complementarios
    // (sin solapamiento de celdas llenas, patrón ART Lab)
    var complementario = allRows[0].every(function(v, i) {
      var v0 = String(v).trim();
      var v1 = String((allRows[1][i] !== undefined ? allRows[1][i] : '')).trim();
      return !(v0 && v1); // No hay posición donde ambas filas tengan valor
    });
    if (complementario) {
      // Fusionar fila 1 y fila 2 como encabezado único (ART Lab)
      headerRow = allRows[0].map(function(v, i) {
        return String(v).trim() || String(allRows[1][i] !== undefined ? allRows[1][i] : '').trim();
      });
      dataStart = 2;
    } else {
      // Fila 1 tiene encabezados reales; fila 2 es la primera fila de datos
      headerRow = allRows[0];
      dataStart = 1;
    }
  } else {
    // Caso normal: fila 1 = encabezados
    headerRow = allRows[0];
    dataStart = 1;
  }

  // Detectar columna Periodo oculta en col A (se excluye de la vista)
  var tienePeriodo = String(headerRow[0]).trim().toLowerCase() === 'periodo';
  var colStart = tienePeriodo ? 1 : 0;
  var headers = headerRow.slice(colStart)
    .map(function(h) { return String(h).trim(); })
    .filter(function(h) { return h !== ''; });

  // Incluir todas las filas no vacías (sin filtro de fechas)
  var dataRowsWithNum = allRows.slice(dataStart)
    .map(function(r, i) { return { data: r, rowNum: i + dataStart + 1 }; })
    .filter(function(item) {
      return item.data.some(function(c) { return String(c).trim() !== ''; });
    });

  var rows = dataRowsWithNum.map(function(item) {
    var r = item.data;
    var obj = { _rowNum: item.rowNum, _periodo: tienePeriodo ? String(r[0]) : '' };
    headers.forEach(function(h, i) {
      obj[h] = r[colStart + i];
      obj[h.toLowerCase()] = r[colStart + i];
    });
    return obj;
  });

  // Filtro por sucursal — solo si la columna existe y el filtro está activo
  if (sucursal && sucursal !== 'Todas') {
    var sucHdrIdx = headers.map(function(h){ return h.toLowerCase(); }).indexOf('sucursal');
    if (sucHdrIdx >= 0) {
      rows = rows.filter(function(row) {
        var val = String(row['Sucursal'] || row['sucursal'] || '').trim();
        return val === '' || val === sucursal; // vacío = hereda todas las sucursales
      });
    }
  }

  return { view: viewId, fuente: nombreHoja, headers: headers, rows: rows,
           updated: new Date().toISOString() };
}

/* ══════════════════════════════════════════════════════════════
   LECTORES INDIVIDUALES
   ══════════════════════════════════════════════════════════════ */
/* Columnas: A=Sucursal | B=Fecha(YYYY-MM-DD) | C=Mes | D=Ingresos | E=Gastos | F=Ciclos | G=CAC | H=Margen */
function readMensual(ss, sheetName, fechaInicio, fechaFin, sucursal) {
  var hoja = ss.getSheetByName(sheetName);
  if (!hoja) return { meses:[], ingresos:[], gastos:[], ciclos:[], cac:[], margen:[] };
  var rows = hoja.getDataRange().getValues();
  if (rows.length < 2) return { meses:[], ingresos:[], gastos:[], ciclos:[], cac:[], margen:[] };
  // Detectar columnas por encabezado (soporta columnas en cualquier orden)
  var hdrs = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iSuc    = hdrs.indexOf('sucursal');
  var iFecha  = hdrs.indexOf('fecha');  if (iFecha  < 0) iFecha  = 1;
  var iMes    = hdrs.indexOf('mes');    if (iMes    < 0) iMes    = 2;
  var iIngr   = hdrs.indexOf('ingresos'); if (iIngr < 0) iIngr   = 3;
  var iGast   = hdrs.indexOf('gastos');   if (iGast < 0) iGast   = 4;
  var iCiclos = hdrs.indexOf('ciclos');   if (iCiclos < 0) iCiclos = 5;
  var iCac    = hdrs.indexOf('cac');      if (iCac  < 0) iCac    = 6;
  var iMargen = hdrs.indexOf('margen');   if (iMargen < 0) iMargen = 7;
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[iFecha]).trim();
    if (f < fechaInicio || f > fechaFin) return false;
    if (sucursal && sucursal !== 'Todas' && iSuc >= 0) {
      var s = String(r[iSuc] || '').trim();
      if (s && s !== sucursal) return false;
    }
    return true;
  });
  data.sort(function(a, b) { return String(a[iFecha]) < String(b[iFecha]) ? -1 : 1; });
  return {
    meses:    data.map(function(r) { return String(r[iMes]); }),
    ingresos: data.map(function(r) { return Number(r[iIngr]); }),
    gastos:   data.map(function(r) { return Number(r[iGast]); }),
    ciclos:   data.map(function(r) { return Number(r[iCiclos]); }),
    cac:      data.map(function(r) { return Number(r[iCac]); }),
    margen:   data.map(function(r) { return Number(r[iMargen]); })
  };
}

/* ── Funciones financieras definidas en api_finance.gs (mismo proyecto GAS) ──
   readCashFlow, readServicios, readFunnel, readAlertas, readDonut,
   readPaisesOrigen, readCostos, readEstadoResultados, readOperatingPL,
   _buildPLReport, readBanksData, saveBankRow, doPost
   ────────────────────────────────────────────────────────────────────────── */
function saveCajaChicaIngreso(body) {
  try {
    var sh = getCajaChicaSheet();
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    var fecha    = String(body.fecha    || '').trim();
    var concepto = String(body.concepto || '').trim();
    var entrada  = parseFloat(body.entrada) || 0;

    if (!fecha || !concepto || !entrada) {
      return { ok: false, error: 'fecha, concepto y entrada son requeridos' };
    }

    // Buscar primera fila vacía (placeholder) o agregar al final
    var targetRow = -1;
    for (var r = 1; r < data.length; r++) {
      if (!String(data[r][iFecha] || '').trim() && !String(data[r][iConcepto] || '').trim()) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow === -1) targetRow = sh.getLastRow() + 1;

    sh.getRange(targetRow, iFecha + 1).setValue(fecha);
    sh.getRange(targetRow, iConcepto + 1).setValue(concepto);
    sh.getRange(targetRow, iEntrada + 1).setValue(entrada);
    SpreadsheetApp.flush();

    var nuevoTotal = iTotal >= 0 ? (sh.getRange(targetRow, iTotal + 1).getValue() || 0) : 0;
    return { ok: true, rowNum: targetRow, saldoFinal: Number(nuevoTotal) };
  } catch(ex) {
    return { ok: false, error: ex.message };
  }
}
