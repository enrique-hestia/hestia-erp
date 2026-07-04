/* ==============================================================
   inventario.gs — Inventario de Medicamentos
   --------------------------------------------------------------
   Reemplaza las hojas-matriz (Procedimientos/Estimulación → columna
   por producto) por un libro de movimientos: un renglón por entrada
   o salida, con saldo corrido calculado por el sistema, no por
   fórmula de Sheets. Diseño generalizable — el campo Modulo permite
   sumar Insumos Qx/Lab más adelante sin rehacer el esquema.

   Vive en un spreadsheet dedicado (MED_INV_SS_ID en config.gs),
   separado del archivo "Inventarios" original (que se conserva
   intacto como respaldo histórico).

   Corre setupInventarioMedicamentos() UNA VEZ desde el editor de
   Apps Script para crear el spreadsheet — copia el ID que imprime
   en el log a MED_INV_SS_ID en config.gs y vuelve a desplegar.
   ============================================================== */

var MEDINV_CAT_TAB   = 'Catalogo_Medicamentos';
var MEDINV_MOV_TAB   = 'Movimientos_Inventario';
var MEDINV_COMBO_TAB = 'Combos';

var MEDINV_CAT_HEADERS = [
  'SKU', 'Nombre', 'Categoria', 'Unidad', 'StockMinimo', 'StockMaximo',
  'CostoUnitario', 'ProveedorPreferido', 'StockActual', 'Activo',
  'FechaAlta', 'UsuarioAlta', 'Notas'
];
var MEDINV_MOV_HEADERS = [
  'ID', 'Fecha', 'Modulo', 'SKU', 'Nombre', 'Tipo', 'Cantidad', 'Motivo',
  'Referencia', 'CostoUnitario', 'SaldoResultante', 'Usuario', 'Notas'
];
var MEDINV_COMBO_HEADERS = [
  'ID', 'ProductoIngresos', 'SKU', 'NombreMedicamento', 'CantidadPorUnidad', 'Activo'
];

var MEDINV_CATEGORIAS = ['Estimulación', 'Analgésicos', 'Anestesia', 'Antibióticos', 'Emergencia', 'Otros'];

/* ── Setup (correr UNA VEZ desde el editor de Apps Script) ───────── */
function setupInventarioMedicamentos() {
  var ss = SpreadsheetApp.create('Hestia ERP - Inventario Medicamentos');

  var catSh = ss.getSheets()[0];
  catSh.setName(MEDINV_CAT_TAB);
  catSh.getRange(1, 1, 1, MEDINV_CAT_HEADERS.length).setValues([MEDINV_CAT_HEADERS]);
  catSh.getRange(1, 1, 1, MEDINV_CAT_HEADERS.length).setFontWeight('bold').setBackground('#c46a7a').setFontColor('#ffffff');
  catSh.setFrozenRows(1);
  var catCatCol = MEDINV_CAT_HEADERS.indexOf('Categoria') + 1;
  catSh.getRange(2, catCatCol, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MEDINV_CATEGORIAS, true).setAllowInvalid(true).build()
  );
  var catActivoCol = MEDINV_CAT_HEADERS.indexOf('Activo') + 1;
  catSh.getRange(2, catActivoCol, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE'], true).setAllowInvalid(true).build()
  );
  catSh.autoResizeColumns(1, MEDINV_CAT_HEADERS.length);

  var movSh = ss.insertSheet(MEDINV_MOV_TAB);
  movSh.getRange(1, 1, 1, MEDINV_MOV_HEADERS.length).setValues([MEDINV_MOV_HEADERS]);
  movSh.getRange(1, 1, 1, MEDINV_MOV_HEADERS.length).setFontWeight('bold').setBackground('#1a252f').setFontColor('#ffffff');
  movSh.setFrozenRows(1);
  var movTipoCol = MEDINV_MOV_HEADERS.indexOf('Tipo') + 1;
  movSh.getRange(2, movTipoCol, 5000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Entrada', 'Salida', 'Ajuste'], true).setAllowInvalid(true).build()
  );
  movSh.autoResizeColumns(1, MEDINV_MOV_HEADERS.length);

  var comboSh = ss.insertSheet(MEDINV_COMBO_TAB);
  comboSh.getRange(1, 1, 1, MEDINV_COMBO_HEADERS.length).setValues([MEDINV_COMBO_HEADERS]);
  comboSh.getRange(1, 1, 1, MEDINV_COMBO_HEADERS.length).setFontWeight('bold').setBackground('#1a252f').setFontColor('#ffffff');
  comboSh.setFrozenRows(1);
  comboSh.autoResizeColumns(1, MEDINV_COMBO_HEADERS.length);

  Logger.log('========== INVENTARIO DE MEDICAMENTOS CREADO ==========');
  Logger.log('Spreadsheet ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('SIGUIENTE PASO: copia ese ID a MED_INV_SS_ID en config.gs y vuelve a desplegar.');
  return { ok: true, id: ss.getId(), url: ss.getUrl() };
}

function _medInvSS() {
  if (!MED_INV_SS_ID) throw new Error('MED_INV_SS_ID no configurado — corre setupInventarioMedicamentos() y pega el ID en config.gs');
  return SpreadsheetApp.openById(MED_INV_SS_ID);
}

function _medInvSheet(tab) {
  var ss = _medInvSS();
  var sh = ss.getSheetByName(tab);
  if (!sh) throw new Error('No se encontró la hoja ' + tab + ' — corre setupInventarioMedicamentos() de nuevo');
  return sh;
}

function _medInvColIdx(headers, name) {
  var i = headers.indexOf(name);
  if (i < 0) throw new Error('Columna ' + name + ' no encontrada');
  return i;
}

/* ── Catálogo ─────────────────────────────────────────────────────── */
function readCatalogoMedicamentos() {
  try {
    var sh = _medInvSheet(MEDINV_CAT_TAB);
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!String(r[0] || '').trim()) continue;
      rows.push({
        _rowNum: i + 1,
        sku: String(r[_medInvColIdx(hdrs, 'SKU')] || ''),
        nombre: String(r[_medInvColIdx(hdrs, 'Nombre')] || ''),
        categoria: String(r[_medInvColIdx(hdrs, 'Categoria')] || ''),
        unidad: String(r[_medInvColIdx(hdrs, 'Unidad')] || ''),
        stockMinimo: Number(r[_medInvColIdx(hdrs, 'StockMinimo')]) || 0,
        stockMaximo: Number(r[_medInvColIdx(hdrs, 'StockMaximo')]) || 0,
        costoUnitario: Number(r[_medInvColIdx(hdrs, 'CostoUnitario')]) || 0,
        proveedorPreferido: String(r[_medInvColIdx(hdrs, 'ProveedorPreferido')] || ''),
        stockActual: Number(r[_medInvColIdx(hdrs, 'StockActual')]) || 0,
        activo: String(r[_medInvColIdx(hdrs, 'Activo')] || '').toUpperCase() !== 'FALSE',
        notas: String(r[_medInvColIdx(hdrs, 'Notas')] || '')
      });
    }
    rows.sort(function (a, b) { return a.nombre.toLowerCase() < b.nombre.toLowerCase() ? -1 : 1; });
    var kpis = {
      total: rows.length,
      bajoMinimo: rows.filter(function (r) { return r.activo && r.stockActual < r.stockMinimo; }).length,
      valorInventario: rows.reduce(function (s, r) { return s + r.stockActual * r.costoUnitario; }, 0)
    };
    return { ok: true, rows: rows, kpis: kpis, categorias: MEDINV_CATEGORIAS };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function _getNextMedSku(sh, hdrs) {
  var data = sh.getDataRange().getValues();
  var skuCol = _medInvColIdx(hdrs, 'SKU');
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var m = String(data[i][skuCol] || '').match(/^MED-(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return 'MED-' + String(max + 1).padStart(4, '0');
}

function saveMedicamento(body) {
  try {
    var nombre = String(body.nombre || '').trim();
    if (!nombre) return { ok: false, error: 'El nombre es obligatorio.' };
    var sh = _medInvSheet(MEDINV_CAT_TAB);
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    var data = sh.getDataRange().getValues();
    var nombreCol = _medInvColIdx(hdrs, 'Nombre');
    var nuevoNorm = nombre.toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][nombreCol] || '').trim().toLowerCase() === nuevoNorm) {
        return { ok: false, error: 'Ya existe un medicamento con ese nombre.' };
      }
    }

    var sku = _getNextMedSku(sh, hdrs);
    var fecha = new Date();
    var fechaStr = fecha.getFullYear() + '-' + String(fecha.getMonth() + 1).padStart(2, '0') + '-' + String(fecha.getDate()).padStart(2, '0');
    var row = new Array(hdrs.length).fill('');
    row[_medInvColIdx(hdrs, 'SKU')] = sku;
    row[_medInvColIdx(hdrs, 'Nombre')] = nombre;
    row[_medInvColIdx(hdrs, 'Categoria')] = String(body.categoria || '');
    row[_medInvColIdx(hdrs, 'Unidad')] = String(body.unidad || '');
    row[_medInvColIdx(hdrs, 'StockMinimo')] = Number(body.stockMinimo) || 0;
    row[_medInvColIdx(hdrs, 'StockMaximo')] = Number(body.stockMaximo) || 0;
    row[_medInvColIdx(hdrs, 'CostoUnitario')] = Number(body.costoUnitario) || 0;
    row[_medInvColIdx(hdrs, 'ProveedorPreferido')] = String(body.proveedorPreferido || '');
    row[_medInvColIdx(hdrs, 'StockActual')] = Number(body.stockInicial) || 0;
    row[_medInvColIdx(hdrs, 'Activo')] = 'TRUE';
    row[_medInvColIdx(hdrs, 'FechaAlta')] = fechaStr;
    row[_medInvColIdx(hdrs, 'UsuarioAlta')] = String(body.usuario || '');
    row[_medInvColIdx(hdrs, 'Notas')] = String(body.notas || '');
    sh.appendRow(row);

    // Si arranca con stock inicial, registrar el movimiento correspondiente
    var stockInicial = Number(body.stockInicial) || 0;
    if (stockInicial > 0) {
      _registrarMovimientoInventario({
        sku: sku, nombre: nombre, tipo: 'Entrada', cantidad: stockInicial,
        motivo: 'Saldo inicial', referencia: '', costoUnitario: Number(body.costoUnitario) || 0,
        usuario: body.usuario || '', modulo: 'Medicamentos', notas: 'Alta de catálogo'
      });
    }
    try { logAudit(body.usuario || '', 'Medicamentos', 'Alta', sku, 'nombre', '', nombre); } catch (e) {}
    return { ok: true, sku: sku };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function updateMedicamento(body) {
  try {
    var rowNum = Number(body.rowNum);
    if (!rowNum) return { ok: false, error: 'Falta rowNum' };
    var sh = _medInvSheet(MEDINV_CAT_TAB);
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var campos = ['nombre', 'categoria', 'unidad', 'stockMinimo', 'stockMaximo', 'costoUnitario', 'proveedorPreferido', 'notas'];
    var colMap = {
      nombre: 'Nombre', categoria: 'Categoria', unidad: 'Unidad', stockMinimo: 'StockMinimo',
      stockMaximo: 'StockMaximo', costoUnitario: 'CostoUnitario', proveedorPreferido: 'ProveedorPreferido', notas: 'Notas'
    };
    campos.forEach(function (f) {
      if (body.hasOwnProperty(f)) {
        sh.getRange(rowNum, _medInvColIdx(hdrs, colMap[f]) + 1).setValue(body[f]);
      }
    });
    if (body.hasOwnProperty('activo')) {
      sh.getRange(rowNum, _medInvColIdx(hdrs, 'Activo') + 1).setValue(body.activo ? 'TRUE' : 'FALSE');
    }
    try { logAudit(body.usuario || '', 'Medicamentos', 'Edicion', body.sku || '', '', '', 'fila ' + rowNum); } catch (e) {}
    return { ok: true };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Motor de movimientos: única puerta de entrada para cambiar stock ──
   Mantiene el saldo en Catalogo_Medicamentos (columna StockActual) y dice
   registro en Movimientos_Inventario con el saldo resultante — el saldo
   nunca se recalcula con SUM(), lo escribe el sistema en cada movimiento. */
function _registrarMovimientoInventario(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var catSh = _medInvSheet(MEDINV_CAT_TAB);
    var catHdrs = catSh.getRange(1, 1, 1, catSh.getLastColumn()).getValues()[0];
    var catData = catSh.getDataRange().getValues();
    var skuCol = _medInvColIdx(catHdrs, 'SKU');
    var stockCol = _medInvColIdx(catHdrs, 'StockActual');
    var rowIdx = -1;
    for (var i = 1; i < catData.length; i++) {
      if (String(catData[i][skuCol] || '').trim() === String(p.sku).trim()) { rowIdx = i; break; }
    }
    if (rowIdx < 0) return { ok: false, error: 'SKU no encontrado en el catálogo: ' + p.sku };

    var saldoActual = Number(catData[rowIdx][stockCol]) || 0;
    var cantidad = Math.abs(Number(p.cantidad) || 0);
    var saldoNuevo;
    if (p.tipo === 'Entrada') saldoNuevo = saldoActual + cantidad;
    else if (p.tipo === 'Salida') saldoNuevo = saldoActual - cantidad;
    else saldoNuevo = Number(p.cantidad); // Ajuste: cantidad es el saldo absoluto deseado

    catSh.getRange(rowIdx + 1, stockCol + 1).setValue(saldoNuevo);

    var movSh = _medInvSheet(MEDINV_MOV_TAB);
    var movHdrs = movSh.getRange(1, 1, 1, movSh.getLastColumn()).getValues()[0];
    var nextId = 'MOV-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    var fecha = p.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var row = new Array(movHdrs.length).fill('');
    row[_medInvColIdx(movHdrs, 'ID')] = nextId;
    row[_medInvColIdx(movHdrs, 'Fecha')] = fecha;
    row[_medInvColIdx(movHdrs, 'Modulo')] = p.modulo || 'Medicamentos';
    row[_medInvColIdx(movHdrs, 'SKU')] = p.sku;
    row[_medInvColIdx(movHdrs, 'Nombre')] = p.nombre || '';
    row[_medInvColIdx(movHdrs, 'Tipo')] = p.tipo;
    row[_medInvColIdx(movHdrs, 'Cantidad')] = p.tipo === 'Ajuste' ? (saldoNuevo - saldoActual) : cantidad;
    row[_medInvColIdx(movHdrs, 'Motivo')] = p.motivo || '';
    row[_medInvColIdx(movHdrs, 'Referencia')] = p.referencia || '';
    row[_medInvColIdx(movHdrs, 'CostoUnitario')] = Number(p.costoUnitario) || 0;
    row[_medInvColIdx(movHdrs, 'SaldoResultante')] = saldoNuevo;
    row[_medInvColIdx(movHdrs, 'Usuario')] = p.usuario || '';
    row[_medInvColIdx(movHdrs, 'Notas')] = p.notas || '';
    movSh.appendRow(row);

    return { ok: true, saldoNuevo: saldoNuevo, movimientoId: nextId };
  } catch (ex) {
    return { ok: false, error: ex.message };
  } finally {
    lock.releaseLock();
  }
}

/* ── Compras (entradas de mercancía) ─────────────────────────────── */
function registrarCompraMedicamento(body) {
  try {
    var sku = String(body.sku || '').trim();
    var cantidad = Number(body.cantidad) || 0;
    if (!sku || cantidad <= 0) return { ok: false, error: 'SKU y cantidad son obligatorios' };

    var catSh = _medInvSheet(MEDINV_CAT_TAB);
    var catHdrs = catSh.getRange(1, 1, 1, catSh.getLastColumn()).getValues()[0];
    var catData = catSh.getDataRange().getValues();
    var skuCol = _medInvColIdx(catHdrs, 'SKU');
    var nombreCol = _medInvColIdx(catHdrs, 'Nombre');
    var costoCol = _medInvColIdx(catHdrs, 'CostoUnitario');
    var nombre = '';
    for (var i = 1; i < catData.length; i++) {
      if (String(catData[i][skuCol] || '').trim() === sku) { nombre = String(catData[i][nombreCol] || ''); break; }
    }
    if (!nombre) return { ok: false, error: 'SKU no encontrado en el catálogo' };

    var costoUnitario = Number(body.costoUnitario) || 0;
    var referencia = [body.proveedor, body.factura].filter(Boolean).join(' — ');
    var resultado = _registrarMovimientoInventario({
      sku: sku, nombre: nombre, tipo: 'Entrada', cantidad: cantidad,
      motivo: 'Compra', referencia: referencia, costoUnitario: costoUnitario,
      usuario: body.usuario || '', modulo: 'Medicamentos', notas: body.notas || ''
    });
    if (!resultado.ok) return resultado;

    // Actualiza el costo unitario vigente en el catálogo con el de la compra más reciente
    if (costoUnitario > 0) {
      for (var j = 1; j < catData.length; j++) {
        if (String(catData[j][skuCol] || '').trim() === sku) {
          catSh.getRange(j + 1, costoCol + 1).setValue(costoUnitario);
          break;
        }
      }
    }
    try { logAudit(body.usuario || '', 'Medicamentos', 'Compra', sku, 'cantidad', '', String(cantidad)); } catch (e) {}
    return { ok: true, saldoNuevo: resultado.saldoNuevo, movimientoId: resultado.movimientoId };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Ajuste manual (merma, caducidad, conteo físico) ─────────────── */
function ajustarInventarioMedicamento(body) {
  try {
    var sku = String(body.sku || '').trim();
    if (!sku) return { ok: false, error: 'Falta SKU' };
    var motivo = String(body.motivo || 'Ajuste Manual');
    var tipo = (motivo === 'Merma' || motivo === 'Caducado') ? 'Salida' : 'Ajuste';
    var cantidad = Number(body.cantidad) || 0;

    var catSh = _medInvSheet(MEDINV_CAT_TAB);
    var catHdrs = catSh.getRange(1, 1, 1, catSh.getLastColumn()).getValues()[0];
    var catData = catSh.getDataRange().getValues();
    var skuCol = _medInvColIdx(catHdrs, 'SKU');
    var nombreCol = _medInvColIdx(catHdrs, 'Nombre');
    var nombre = '';
    for (var i = 1; i < catData.length; i++) {
      if (String(catData[i][skuCol] || '').trim() === sku) { nombre = String(catData[i][nombreCol] || ''); break; }
    }
    if (!nombre) return { ok: false, error: 'SKU no encontrado en el catálogo' };

    return _registrarMovimientoInventario({
      sku: sku, nombre: nombre, tipo: tipo, cantidad: cantidad, motivo: motivo,
      referencia: body.referencia || '', usuario: body.usuario || '', modulo: 'Medicamentos', notas: body.notas || ''
    });
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Historial de movimientos (auditoría) ────────────────────────── */
function readMovimientosInventario(filtros) {
  try {
    filtros = filtros || {};
    var sh = _medInvSheet(MEDINV_MOV_TAB);
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var rows = [];
    for (var i = data.length - 1; i >= 1; i--) { // más reciente primero
      var r = data[i];
      if (!String(r[0] || '').trim()) continue;
      var obj = {
        id: String(r[_medInvColIdx(hdrs, 'ID')] || ''),
        fecha: String(r[_medInvColIdx(hdrs, 'Fecha')] || ''),
        modulo: String(r[_medInvColIdx(hdrs, 'Modulo')] || ''),
        sku: String(r[_medInvColIdx(hdrs, 'SKU')] || ''),
        nombre: String(r[_medInvColIdx(hdrs, 'Nombre')] || ''),
        tipo: String(r[_medInvColIdx(hdrs, 'Tipo')] || ''),
        cantidad: Number(r[_medInvColIdx(hdrs, 'Cantidad')]) || 0,
        motivo: String(r[_medInvColIdx(hdrs, 'Motivo')] || ''),
        referencia: String(r[_medInvColIdx(hdrs, 'Referencia')] || ''),
        saldoResultante: Number(r[_medInvColIdx(hdrs, 'SaldoResultante')]) || 0,
        usuario: String(r[_medInvColIdx(hdrs, 'Usuario')] || '')
      };
      if (filtros.sku && obj.sku !== filtros.sku) continue;
      if (filtros.fechaInicio && obj.fecha < filtros.fechaInicio) continue;
      if (filtros.fechaFin && obj.fecha > filtros.fechaFin) continue;
      rows.push(obj);
      if (rows.length >= 500) break; // tope razonable para no saturar la respuesta
    }
    return { ok: true, rows: rows };
  } catch (ex) { return { ok: false, error: ex.message, rows: [] }; }
}

/* ── Combos: qué medicamentos descuenta un Producto de Ingresos ──── */
function readCombos() {
  try {
    var sh = _medInvSheet(MEDINV_COMBO_TAB);
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!String(r[_medInvColIdx(hdrs, 'ProductoIngresos')] || '').trim()) continue;
      rows.push({
        _rowNum: i + 1,
        id: String(r[_medInvColIdx(hdrs, 'ID')] || ''),
        productoIngresos: String(r[_medInvColIdx(hdrs, 'ProductoIngresos')] || ''),
        sku: String(r[_medInvColIdx(hdrs, 'SKU')] || ''),
        nombreMedicamento: String(r[_medInvColIdx(hdrs, 'NombreMedicamento')] || ''),
        cantidadPorUnidad: Number(r[_medInvColIdx(hdrs, 'CantidadPorUnidad')]) || 0,
        activo: String(r[_medInvColIdx(hdrs, 'Activo')] || '').toUpperCase() !== 'FALSE'
      });
    }
    return { ok: true, rows: rows };
  } catch (ex) { return { ok: false, error: ex.message, rows: [] }; }
}

function saveCombo(body) {
  try {
    var producto = String(body.productoIngresos || '').trim();
    var sku = String(body.sku || '').trim();
    var cantidad = Number(body.cantidadPorUnidad) || 0;
    if (!producto || !sku || cantidad <= 0) return { ok: false, error: 'Producto, medicamento y cantidad son obligatorios' };

    var catSh = _medInvSheet(MEDINV_CAT_TAB);
    var catData = catSh.getDataRange().getValues();
    var catHdrs = catData[0];
    var skuCol = _medInvColIdx(catHdrs, 'SKU');
    var nombreCol = _medInvColIdx(catHdrs, 'Nombre');
    var nombreMed = '';
    for (var i = 1; i < catData.length; i++) {
      if (String(catData[i][skuCol] || '').trim() === sku) { nombreMed = String(catData[i][nombreCol] || ''); break; }
    }
    if (!nombreMed) return { ok: false, error: 'Medicamento no encontrado en el catálogo' };

    var sh = _medInvSheet(MEDINV_COMBO_TAB);
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var id = 'COMBO-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    var row = new Array(hdrs.length).fill('');
    row[_medInvColIdx(hdrs, 'ID')] = id;
    row[_medInvColIdx(hdrs, 'ProductoIngresos')] = producto;
    row[_medInvColIdx(hdrs, 'SKU')] = sku;
    row[_medInvColIdx(hdrs, 'NombreMedicamento')] = nombreMed;
    row[_medInvColIdx(hdrs, 'CantidadPorUnidad')] = cantidad;
    row[_medInvColIdx(hdrs, 'Activo')] = 'TRUE';
    sh.appendRow(row);
    try { logAudit(body.usuario || '', 'Medicamentos', 'ComboAlta', id, '', '', producto + ' -> ' + nombreMed); } catch (e) {}
    return { ok: true, id: id };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function eliminarCombo(body) {
  try {
    var rowNum = Number(body.rowNum);
    if (!rowNum) return { ok: false, error: 'Falta rowNum' };
    var sh = _medInvSheet(MEDINV_COMBO_TAB);
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    sh.getRange(rowNum, _medInvColIdx(hdrs, 'Activo') + 1).setValue('FALSE');
    return { ok: true };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Corrige el menú (hoja "Menu" en SHEET_ID) ────────────────────
   El nav "Medicamentos" (prod-meds) ya existe como vista suelta, y
   sus hijos "Orden de Compra"/"Estimulación" quedaron huérfanos
   (Padre="medicamentos", pero el ID real del grupo es "prod-meds")
   apuntando al sistema viejo. Se convierte prod-meds en un grupo
   (igual que Quirofano/Laboratorio) con 4 hijos del sistema nuevo:
   Catálogo, Compras, Movimientos, Combos. Correr UNA VEZ desde el
   editor de Apps Script. */
function configurarMenuMedicamentos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Menu');
  if (!sh) return { ok: false, error: 'No se encontró la hoja Menu' };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var idxCol = hdrs.indexOf('ID');
  var tipoCol = hdrs.indexOf('Tipo');
  var fuenteCol = hdrs.indexOf('Fuente');

  // Quitar los hijos huérfanos del sistema viejo
  var filasABorrar = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][idxCol] || '').trim();
    if (id === 'med-compra' || id === 'med-estimulacion') filasABorrar.push(i + 1);
    if (id === 'prod-meds') {
      sh.getRange(i + 1, tipoCol + 1).setValue('grupo');
      sh.getRange(i + 1, fuenteCol + 1).setValue('');
    }
  }
  filasABorrar.sort(function (a, b) { return b - a; }); // de abajo hacia arriba
  filasABorrar.forEach(function (rowNum) { sh.deleteRow(rowNum); });

  // Agregar los 4 hijos del sistema nuevo bajo prod-meds. IDs elegidos para NO
  // contener 'catalogo'/'medicamento'/'prod-med' como substring — navigateTo()
  // redirige cualquier viewId que contenga esas palabras a la vista genérica
  // compartida de Productos (view-productos-catalogo), y esta pantalla es un
  // sistema aparte, no ese catálogo de servicios.
  var nuevasFilas = [
    ['meds-lista', 'prod-meds', 'Catálogo', '', 'pill', 1, 'vista', 'meds-lista', 'TRUE'],
    ['meds-compras', 'prod-meds', 'Compras', '', 'shopping-cart', 2, 'vista', 'meds-compras', 'TRUE'],
    ['meds-movimientos', 'prod-meds', 'Movimientos', '', 'history', 3, 'vista', 'meds-movimientos', 'TRUE'],
    ['meds-combos', 'prod-meds', 'Combos', '', 'boxes', 4, 'vista', 'meds-combos', 'TRUE']
  ];
  sh.getRange(sh.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);

  return { ok: true, filasBorradas: filasABorrar.length, filasAgregadas: nuevasFilas.length };
}

/* ── Descuento automático al vender ───────────────────────────────
   Se llama desde saveIngreso() en finance.gs después de guardar la
   venta — nunca debe tumbar la venta si el inventario falla, por eso
   quien la invoca la envuelve en try/catch y solo registra el error. */
function _descontarInventarioPorVenta(opId, lineas, usuario) {
  var combosData = readCombos();
  if (!combosData.ok || !combosData.rows.length) return { ok: true, descontados: 0 };
  var combosPorProducto = {};
  combosData.rows.forEach(function (c) {
    if (!c.activo) return;
    var key = c.productoIngresos.trim().toLowerCase();
    if (!combosPorProducto[key]) combosPorProducto[key] = [];
    combosPorProducto[key].push(c);
  });

  var resultados = [];
  lineas.forEach(function (l) {
    var key = String(l.producto || '').trim().toLowerCase();
    var combos = combosPorProducto[key];
    if (!combos || !combos.length) return;
    var cantLinea = Number(l.cantidad) || 1;
    combos.forEach(function (c) {
      var cantidadADescontar = c.cantidadPorUnidad * cantLinea;
      var r = _registrarMovimientoInventario({
        sku: c.sku, nombre: c.nombreMedicamento, tipo: 'Salida', cantidad: cantidadADescontar,
        motivo: 'Venta', referencia: opId, usuario: usuario || '', modulo: 'Medicamentos',
        notas: 'Producto: ' + l.producto
      });
      resultados.push({ sku: c.sku, cantidad: cantidadADescontar, ok: r.ok, saldoNuevo: r.saldoNuevo, error: r.error });
    });
  });
  return { ok: true, descontados: resultados.length, detalle: resultados };
}
