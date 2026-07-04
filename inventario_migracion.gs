/* ==============================================================
   inventario_migracion.gs — Migración de historial al nuevo
   sistema de Inventario de Medicamentos
   --------------------------------------------------------------
   Lee el spreadsheet "Inventarios" original (MED_SS_ID) — hojas
   Estimulación, Estimulación 2025, Ent. Med, Lista Med — SIN
   modificarlo, y arma los movimientos históricos correspondientes
   en el nuevo sistema (MED_INV_SS_ID).

   Corre primero SIN confirmar=true (modo reporte): regresa cuántos
   medicamentos y movimientos va a crear, para revisar antes de
   escribir nada. Cuando se ve bien, se llama con confirmar=true
   desde el frontend (o directo en el editor) para aplicar.

   No toca Procedimientos / Inventario QX (eso es Insumos Qx, fuera
   de este alcance) ni escribe nada en el spreadsheet original.
   ============================================================== */

function _migInvLeerEstimulacion(ssOrig) {
  var sh = ssOrig.getSheetByName('Estimulación');
  if (!sh) return { medCols: [], movimientos: [] };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0].map(function (h) { return String(h).trim(); });
  var canceladoIdx = -1;
  for (var c = 0; c < hdrs.length; c++) {
    if (hdrs[c].toLowerCase() === 'cancelado') { canceladoIdx = c; break; }
  }
  var medCols = canceladoIdx >= 0 ? hdrs.slice(canceladoIdx + 1).filter(function (h) { return h; }) : [];
  var fechaIdx = hdrs.indexOf('Fecha');
  var pacIdx = hdrs.indexOf('Paciente');

  var movimientos = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var paciente = String(r[pacIdx] || '').trim();
    if (!paciente) continue;
    var fechaVal = r[fechaIdx];
    var fecha = fechaVal instanceof Date
      ? Utilities.formatDate(fechaVal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fechaVal || '').slice(0, 10);
    medCols.forEach(function (medNombre) {
      var col = hdrs.indexOf(medNombre);
      var cant = Number(r[col]) || 0;
      if (cant > 0) {
        movimientos.push({ fecha: fecha, nombre: medNombre, tipo: 'Salida', cantidad: cant, motivo: 'Consumo histórico', referencia: paciente, origen: 'Estimulación' });
      }
    });
  }
  return { medCols: medCols, movimientos: movimientos };
}

function _migInvLeerEstimulacion2025(ssOrig) {
  var sh = ssOrig.getSheetByName('Estimulación 2025');
  if (!sh) return { medCols: [], movimientos: [] };
  var data = sh.getDataRange().getValues();
  var hdrs = data[1].map(function (h) { return String(h).trim(); }); // fila 2 = header real
  var cicloIdx = -1;
  for (var c = 0; c < hdrs.length; c++) {
    if (hdrs[c].toLowerCase() === 'ciclo') { cicloIdx = c; break; }
  }
  var medCols = cicloIdx >= 0 ? hdrs.slice(cicloIdx + 1).filter(function (h) { return h; }) : [];
  var fechaIdx = hdrs.indexOf('Fecha');
  var pacIdx = hdrs.indexOf('Paciente');

  var movimientos = [];
  for (var i = 2; i < data.length; i++) { // datos desde la fila 3 (índice 2)
    var r = data[i];
    var paciente = String(r[pacIdx] || '').trim();
    if (!paciente) continue;
    var fechaVal = r[fechaIdx];
    var fecha = fechaVal instanceof Date
      ? Utilities.formatDate(fechaVal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fechaVal || '').slice(0, 10);
    medCols.forEach(function (medNombre) {
      var col = hdrs.indexOf(medNombre);
      var cant = Number(r[col]) || 0;
      if (cant > 0) {
        movimientos.push({ fecha: fecha, nombre: medNombre, tipo: 'Salida', cantidad: cant, motivo: 'Consumo histórico', referencia: paciente, origen: 'Estimulación 2025' });
      }
    });
  }
  return { medCols: medCols, movimientos: movimientos };
}

function _migInvLeerEntMed(ssOrig) {
  var sh = ssOrig.getSheetByName('Ent. Med');
  if (!sh) return { medNombres: [], movimientos: [] };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0].map(function (h) { return String(h).trim(); });
  var fechaIdx = hdrs.indexOf('Fecha');
  var medIdx = hdrs.indexOf('Medicamento');
  var cantIdx = hdrs.indexOf('Cantidad');
  var precioIdx = hdrs.indexOf('Precio_Unitario');
  var provIdx = hdrs.indexOf('Proveedor');
  var facturaIdx = hdrs.indexOf('Factura');
  var ordenIdx = hdrs.indexOf('No. Orden');

  var medNombres = {};
  var movimientos = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var nombre = String(r[medIdx] || '').trim();
    if (!nombre) continue;
    medNombres[nombre] = true;
    var fechaVal = r[fechaIdx];
    var fecha = fechaVal instanceof Date
      ? Utilities.formatDate(fechaVal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fechaVal || '').slice(0, 10);
    var cantidad = Number(r[cantIdx]) || 0;
    if (cantidad <= 0) continue;
    var referencia = [String(r[ordenIdx] || ''), String(r[provIdx] || ''), String(r[facturaIdx] || '')].filter(Boolean).join(' — ');
    movimientos.push({
      fecha: fecha, nombre: nombre, tipo: 'Entrada', cantidad: cantidad,
      motivo: 'Compra histórica', referencia: referencia, costoUnitario: Number(r[precioIdx]) || 0, origen: 'Ent. Med'
    });
  }
  return { medNombres: Object.keys(medNombres), movimientos: movimientos };
}

function _migInvLeerListaMed(ssOrig) {
  var sh = ssOrig.getSheetByName('Lista Med');
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  var costos = {};
  for (var i = 1; i < data.length; i++) {
    var nombre = String(data[i][0] || '').trim();
    if (!nombre) continue;
    var costo = Number(data[i][1]) || 0; // col B = Costo vigente
    if (costo > 0) costos[nombre] = costo;
  }
  return costos;
}

// Modo reporte (confirmar=false/omitido): arma todo en memoria y regresa un
// resumen para revisar SIN escribir nada. Modo aplicar (confirmar=true):
// hace lo mismo y además escribe el catálogo + los movimientos en bloque.
function migrarHistorialInventario(confirmar) {
  try {
    if (!MED_SS_ID) return { ok: false, error: 'MED_SS_ID no configurado' };
    var ssOrig = SpreadsheetApp.openById(MED_SS_ID);

    var est = _migInvLeerEstimulacion(ssOrig);
    var est2 = _migInvLeerEstimulacion2025(ssOrig);
    var ent = _migInvLeerEntMed(ssOrig);
    var costos = _migInvLeerListaMed(ssOrig);

    // Unión de nombres de medicamento (match EXACTO de texto — variantes con
    // nombre distinto, ej. "Gonal 300" vs "Gonal F 300", quedan como
    // productos separados a propósito; se revisan y fusionan a mano si hace
    // falta, mejor que arriesgar mezclar dos dosis distintas por error).
    var nombresSet = {};
    est.medCols.forEach(function (n) { nombresSet[n] = true; });
    est2.medCols.forEach(function (n) { nombresSet[n] = true; });
    ent.medNombres.forEach(function (n) { nombresSet[n] = true; });
    var nombres = Object.keys(nombresSet).sort();

    // Todos los movimientos juntos, ordenados por fecha para que el saldo
    // corrido tenga sentido cronológico en el historial.
    var todosMovimientos = [].concat(est.movimientos, est2.movimientos, ent.movimientos);
    todosMovimientos.sort(function (a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });

    var resumen = {
      medicamentosEncontrados: nombres.length,
      nombres: nombres,
      movimientosEntrada: todosMovimientos.filter(function (m) { return m.tipo === 'Entrada'; }).length,
      movimientosSalida: todosMovimientos.filter(function (m) { return m.tipo === 'Salida'; }).length,
      totalMovimientos: todosMovimientos.length,
      rangoFechas: todosMovimientos.length ? (todosMovimientos[0].fecha + ' a ' + todosMovimientos[todosMovimientos.length - 1].fecha) : ''
    };

    if (!confirmar) {
      return { ok: true, modo: 'reporte', resumen: resumen };
    }

    // ── Modo aplicar ──────────────────────────────────────────────
    var catSh = _medInvSheet(MEDINV_CAT_TAB);
    var catData = catSh.getDataRange().getValues();
    var catHdrs = catData[0];
    var skuCol = _medInvColIdx(catHdrs, 'SKU');
    var nombreCol = _medInvColIdx(catHdrs, 'Nombre');
    var existentes = {};
    for (var i = 1; i < catData.length; i++) {
      var n = String(catData[i][nombreCol] || '').trim();
      if (n) existentes[n] = String(catData[i][skuCol] || '').trim();
    }

    // Crear en el catálogo los medicamentos que falten
    var maxNum = 0;
    for (var k in existentes) {
      var m = existentes[k].match(/^MED-(\d+)$/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    }
    var nuevasFilas = [];
    var skuPorNombre = {};
    nombres.forEach(function (nombre) {
      if (existentes[nombre]) { skuPorNombre[nombre] = existentes[nombre]; return; }
      maxNum++;
      var sku = 'MED-' + String(maxNum).padStart(4, '0');
      skuPorNombre[nombre] = sku;
      var row = new Array(catHdrs.length).fill('');
      row[_medInvColIdx(catHdrs, 'SKU')] = sku;
      row[_medInvColIdx(catHdrs, 'Nombre')] = nombre;
      row[_medInvColIdx(catHdrs, 'Categoria')] = 'Estimulación';
      row[_medInvColIdx(catHdrs, 'CostoUnitario')] = costos[nombre] || 0;
      row[_medInvColIdx(catHdrs, 'StockActual')] = 0;
      row[_medInvColIdx(catHdrs, 'Activo')] = 'TRUE';
      row[_medInvColIdx(catHdrs, 'FechaAlta')] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      row[_medInvColIdx(catHdrs, 'UsuarioAlta')] = 'migracion';
      row[_medInvColIdx(catHdrs, 'Notas')] = 'Importado de ' + MED_SS_ID;
      nuevasFilas.push(row);
    });
    if (nuevasFilas.length) {
      catSh.getRange(catSh.getLastRow() + 1, 1, nuevasFilas.length, catHdrs.length).setValues(nuevasFilas);
    }

    // Saldo corrido en memoria por SKU, procesando movimientos en orden cronológico
    var saldoPorSku = {};
    var movRows = [];
    var movSh = _medInvSheet(MEDINV_MOV_TAB);
    var movHdrs = movSh.getRange(1, 1, 1, movSh.getLastColumn()).getValues()[0];

    todosMovimientos.forEach(function (m, idx) {
      var sku = skuPorNombre[m.nombre];
      if (!sku) return;
      var saldoActual = saldoPorSku[sku] || 0;
      var saldoNuevo = m.tipo === 'Entrada' ? saldoActual + m.cantidad : saldoActual - m.cantidad;
      saldoPorSku[sku] = saldoNuevo;

      var row = new Array(movHdrs.length).fill('');
      row[_medInvColIdx(movHdrs, 'ID')] = 'MOV-MIG-' + idx;
      row[_medInvColIdx(movHdrs, 'Fecha')] = m.fecha;
      row[_medInvColIdx(movHdrs, 'Modulo')] = 'Medicamentos';
      row[_medInvColIdx(movHdrs, 'SKU')] = sku;
      row[_medInvColIdx(movHdrs, 'Nombre')] = m.nombre;
      row[_medInvColIdx(movHdrs, 'Tipo')] = m.tipo;
      row[_medInvColIdx(movHdrs, 'Cantidad')] = m.cantidad;
      row[_medInvColIdx(movHdrs, 'Motivo')] = m.motivo;
      row[_medInvColIdx(movHdrs, 'Referencia')] = m.referencia || '';
      row[_medInvColIdx(movHdrs, 'CostoUnitario')] = m.costoUnitario || 0;
      row[_medInvColIdx(movHdrs, 'SaldoResultante')] = saldoNuevo;
      row[_medInvColIdx(movHdrs, 'Usuario')] = 'migracion';
      row[_medInvColIdx(movHdrs, 'Notas')] = 'Origen: ' + m.origen;
      movRows.push(row);
    });

    if (movRows.length) {
      movSh.getRange(movSh.getLastRow() + 1, 1, movRows.length, movHdrs.length).setValues(movRows);
    }

    // Escribir el saldo final en el catálogo (StockActual)
    var catData2 = catSh.getDataRange().getValues();
    var stockCol = _medInvColIdx(catHdrs, 'StockActual');
    for (var ri = 1; ri < catData2.length; ri++) {
      var skuRow = String(catData2[ri][skuCol] || '').trim();
      if (saldoPorSku.hasOwnProperty(skuRow)) {
        catSh.getRange(ri + 1, stockCol + 1).setValue(saldoPorSku[skuRow]);
      }
    }

    try { logAudit('migracion', 'Medicamentos', 'MigracionHistorica', '', '', '', movRows.length + ' movimientos importados'); } catch (e) {}

    return { ok: true, modo: 'aplicado', resumen: resumen, medicamentosCreados: nuevasFilas.length, movimientosCreados: movRows.length };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
