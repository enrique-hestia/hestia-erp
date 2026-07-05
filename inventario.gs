/* ==============================================================
   inventario.gs — Inventario (Medicamentos y cualquier producto
   marcado "Inventariable")
   --------------------------------------------------------------
   Un solo catálogo: el stock vive directamente en BD_Productos
   (columnas Inventariable/Unidad/Stock…, ver _bdProdEnsureInventarioCols
   en finance.gs). Esta hoja ya NO mantiene su propio catálogo — solo
   el ledger de movimientos (un renglón por entrada/salida/ajuste, con
   saldo corrido calculado por el sistema, no por fórmula de Sheets) y
   la tabla de Combos, ambos en MED_INV_SS_ID. El campo Modulo del
   ledger guarda la Categoría del producto afectado, para poder filtrar
   por Medicamentos/Insumos/etc. sin necesitar catálogos separados.

   Cualquier producto de BD_Productos puede entrar al sistema: se marca
   "Medicamento" y se activa solo, o se marca "Inventariable" manual
   desde Catálogo General para insumos/reactivos/lo que sea — y con eso
   ya se puede usar como componente de un Combo.
   ============================================================== */

var MEDINV_MOV_TAB   = 'Movimientos_Inventario';
var MEDINV_COMBO_TAB = 'Combos';

var MEDINV_MOV_HEADERS = [
  'ID', 'Fecha', 'Modulo', 'SKU', 'Nombre', 'Tipo', 'Cantidad', 'Motivo',
  'Referencia', 'CostoUnitario', 'SaldoResultante', 'Usuario', 'Notas'
];
var MEDINV_COMBO_HEADERS = [
  'ID', 'ProductoIngresos', 'SKU', 'NombreMedicamento', 'CantidadPorUnidad', 'Activo'
];

// Subcategorías sugeridas para medicamentos (campo Tipo de BD_Productos) —
// no son la Categoria del catálogo general, solo agrupan el tipo de fármaco.
var MEDINV_SUBCATEGORIAS_SUGERIDAS = ['Estimulación', 'Analgésicos', 'Anestesia', 'Antibióticos', 'Emergencia', 'Otros'];

/* ── Setup (correr UNA VEZ desde el editor de Apps Script) ───────── */
function setupInventarioMedicamentos() {
  var ss = SpreadsheetApp.create('Hestia ERP - Inventario Medicamentos');

  var movSh = ss.getSheets()[0];
  movSh.setName(MEDINV_MOV_TAB);
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

  Logger.log('========== INVENTARIO CREADO ==========');
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

/* ── Backfill: activa Inventariable en los productos "Medicamento" que
   ya existían en BD_Productos antes de este sistema, para que aparezcan
   de inmediato en Medicamentos > Catálogo sin tener que volver a
   guardarlos uno por uno. Genera SKU si falta (lo necesitan Combos y
   Movimientos para identificar el producto) y solo inicializa
   StockActual en 0 si estaba vacío — nunca pisa un stock ya capturado.
   Correr UNA VEZ desde el editor de Apps Script. */
function activarInventarioMedicamentosExistentes() {
  var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
  var sh = ss.getSheetByName('BD_Productos');
  if (!sh) return { ok: false, error: 'BD_Productos no encontrada' };
  var cols = _bdProdEnsureInventarioCols(sh);
  var data = sh.getDataRange().getValues();
  var actualizados = 0, skuGenerados = 0;
  for (var i = 1; i < data.length; i++) {
    var categoria = String(data[i][3] || '').trim().toLowerCase();
    if (categoria !== 'medicamento') continue;
    var rowNum = i + 1;
    var yaInventariable = data[i][cols.Inventariable - 1] === true || String(data[i][cols.Inventariable - 1]).toUpperCase() === 'TRUE';
    if (!yaInventariable) { sh.getRange(rowNum, cols.Inventariable).setValue(true); actualizados++; }
    if (!String(data[i][1] || '').trim()) {
      sh.getRange(rowNum, 2).setValue(_getNextSkuConPrefijo(sh, 'MED'));
      skuGenerados++;
    }
    var stockActualCell = data[i][cols.StockActual - 1];
    if (stockActualCell === '' || stockActualCell === null || stockActualCell === undefined) {
      sh.getRange(rowNum, cols.StockActual).setValue(0);
    }
  }
  return { ok: true, actualizados: actualizados, skuGenerados: skuGenerados };
}

/* ── Buscar un producto inventariable en BD_Productos por SKU ─────── */
function _bdProdFindBySku(sku) {
  var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
  var sh = ss.getSheetByName('BD_Productos');
  if (!sh) throw new Error('BD_Productos no encontrada');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() === String(sku).trim()) {
      return {
        sh: sh, rowNum: i + 1,
        productoId: String(data[i][0] || ''), sku: String(data[i][1] || ''),
        descripcion: String(data[i][2] || ''), categoria: String(data[i][3] || '')
      };
    }
  }
  return null;
}

/* ── Catálogo para los dropdowns de Inventario: TODO el Catálogo
   General (no solo lo Inventariable) — al usar un producto en una
   orden de compra, combo o movimiento, el sistema lo marca
   Inventariable automáticamente, sin bloquear la captura. Los KPIs
   de stock sí se calculan solo sobre lo ya inventariable. ─────────── */
function readCatalogoMedicamentos() {
  try {
    var data = readProductos();
    if (!data.ok) return data;
    var rows = data.todosProductos.map(function (p) {
      return {
        productoId: p.id, sku: p.sku, nombre: p.descripcion, categoria: p.categoria,
        tipo: p.tipo, unidad: p.unidad, stockMinimo: p.stockMinimo, stockMaximo: p.stockMaximo,
        costoUnitario: p.costoUnitario, proveedorPreferido: p.proveedorPreferido,
        stockActual: p.stockActual, activo: p.activo, notas: p.notas,
        inventariable: !!p.inventariable
      };
    });
    rows.sort(function (a, b) { return a.nombre.toLowerCase() < b.nombre.toLowerCase() ? -1 : 1; });
    var inv = rows.filter(function (r) { return r.inventariable; });
    var kpis = {
      total: inv.length,
      bajoMinimo: inv.filter(function (r) { return r.activo && r.stockActual < r.stockMinimo; }).length,
      valorInventario: inv.reduce(function (s, r) { return s + r.stockActual * r.costoUnitario; }, 0)
    };
    // "categorias" aquí son las SUBcategorías de medicamento (Estimulación,
    // Analgésicos…), que viven en el campo Tipo de BD_Productos — no las
    // categorías generales del catálogo (Medicamento/Tratamientos/etc). Se
    // arma con los valores de Tipo ya usados + la lista sugerida de base,
    // para que el picker no se quede vacío en un catálogo nuevo.
    var tiposUsados = {};
    rows.forEach(function (r) { if (r.categoria === 'Medicamento' && r.tipo) tiposUsados[r.tipo] = true; });
    MEDINV_SUBCATEGORIAS_SUGERIDAS.forEach(function (c) { tiposUsados[c] = true; });
    var categorias = Object.keys(tiposUsados).sort();
    return { ok: true, rows: rows, kpis: kpis, categorias: categorias };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Alta / edición: delegan en el catálogo único (BD_Productos) ──── */
function saveMedicamento(body) {
  body.descripcion = body.descripcion || body.nombre;
  body.categoria = body.categoria || 'Medicamento';
  body.inventariable = true;
  if (!String(body.sku || '').trim()) {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var prodSheet = ss.getSheetByName('BD_Productos');
    if (!prodSheet) { setupBDProductos(); prodSheet = ss.getSheetByName('BD_Productos'); }
    body.sku = _getNextSkuConPrefijo(prodSheet, 'MED');
  }
  return saveNewProducto(body);
}

function updateMedicamento(body) {
  if (body.nombre !== undefined && body.descripcion === undefined) body.descripcion = body.nombre;
  body.inventariable = true;
  return updateProducto(body);
}

/* ── Motor de movimientos: única puerta de entrada para cambiar stock ──
   El saldo vive en BD_Productos (columna StockActual) y cada cambio se
   registra en Movimientos_Inventario con el saldo resultante — el saldo
   nunca se recalcula con SUM(), lo escribe el sistema en cada movimiento. */
function _registrarMovimientoInventario(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var prod = _bdProdFindBySku(p.sku);
    if (!prod) return { ok: false, error: 'SKU no encontrado en el catálogo: ' + p.sku };

    var cols = _bdProdEnsureInventarioCols(prod.sh);
    var saldoActual = Number(prod.sh.getRange(prod.rowNum, cols.StockActual).getValue()) || 0;
    var cantidad = Math.abs(Number(p.cantidad) || 0);
    var saldoNuevo;
    if (p.tipo === 'Entrada') saldoNuevo = saldoActual + cantidad;
    else if (p.tipo === 'Salida') saldoNuevo = saldoActual - cantidad;
    else saldoNuevo = Number(p.cantidad); // Ajuste: cantidad es el saldo absoluto deseado

    prod.sh.getRange(prod.rowNum, cols.StockActual).setValue(saldoNuevo);
    prod.sh.getRange(prod.rowNum, cols.Inventariable).setValue(true);

    var movSh = _medInvSheet(MEDINV_MOV_TAB);
    var movHdrs = movSh.getRange(1, 1, 1, movSh.getLastColumn()).getValues()[0];
    var nextId = 'MOV-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    var fecha = p.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var row = new Array(movHdrs.length).fill('');
    row[_medInvColIdx(movHdrs, 'ID')] = nextId;
    row[_medInvColIdx(movHdrs, 'Fecha')] = fecha;
    row[_medInvColIdx(movHdrs, 'Modulo')] = p.modulo || prod.categoria || 'Productos';
    row[_medInvColIdx(movHdrs, 'SKU')] = p.sku;
    row[_medInvColIdx(movHdrs, 'Nombre')] = p.nombre || prod.descripcion || '';
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

/* ── Ajuste manual (merma, caducidad, conteo físico) ─────────────── */
function ajustarInventarioMedicamento(body) {
  try {
    var sku = String(body.sku || '').trim();
    if (!sku) return { ok: false, error: 'Falta SKU' };
    var motivo = String(body.motivo || 'Ajuste Manual');
    var tipo = (motivo === 'Merma' || motivo === 'Caducado') ? 'Salida' : 'Ajuste';
    var cantidad = Number(body.cantidad) || 0;

    var prod = _bdProdFindBySku(sku);
    if (!prod) return { ok: false, error: 'SKU no encontrado en el catálogo' };

    return _registrarMovimientoInventario({
      sku: sku, nombre: prod.descripcion, tipo: tipo, cantidad: cantidad, motivo: motivo,
      referencia: body.referencia || '', usuario: body.usuario || '', modulo: prod.categoria, notas: body.notas || ''
    });
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Sobrante: lo que no se usó de un componente al aplicar un combo
   (ej. sobra medio frasco) regresa al stock en tiempo real. Acción
   rápida desde Movimientos — no se mete en la captura de Ingresos para
   no hacerla más lenta durante la consulta. ─────────────────────── */
function registrarSobranteInventario(body) {
  try {
    var sku = String(body.sku || '').trim();
    var cantidad = Number(body.cantidad) || 0;
    if (!sku || cantidad <= 0) return { ok: false, error: 'SKU y cantidad son obligatorios' };
    var prod = _bdProdFindBySku(sku);
    if (!prod) return { ok: false, error: 'SKU no encontrado en el catálogo' };

    return _registrarMovimientoInventario({
      sku: sku, nombre: prod.descripcion, tipo: 'Entrada', cantidad: cantidad, motivo: 'Sobrante',
      referencia: String(body.referencia || ''), usuario: body.usuario || '', modulo: prod.categoria,
      notas: body.notas || ''
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

/* ── Combos: qué productos inventariables descuenta un Producto de
   Ingresos — el componente puede ser cualquier producto marcado
   Inventariable (medicamento, insumo, reactivo…), no solo medicamentos,
   para poder armar tratamientos compuestos (ej. Estimulación Ovárica
   Controlada = medicamentos + insumos + serologías). ──────────────── */
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
    if (!producto || !sku || cantidad <= 0) return { ok: false, error: 'Producto, componente y cantidad son obligatorios' };

    var full = null;
    var data = readProductos();
    if (data.ok) full = data.todosProductos.filter(function (p) { return p.sku === sku; })[0];
    if (!full) return { ok: false, error: 'Componente no encontrado en el catálogo' };
    // Si el componente aún no está marcado Inventariable, se activa aquí
    // mismo — usarlo en un combo implica que su stock debe controlarse.
    if (!full.inventariable) {
      var prodRef = _bdProdFindBySku(sku);
      if (prodRef) {
        var invCols = _bdProdEnsureInventarioCols(prodRef.sh);
        prodRef.sh.getRange(prodRef.rowNum, invCols.Inventariable).setValue(true);
      }
    }

    var sh = _medInvSheet(MEDINV_COMBO_TAB);
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var id = 'COMBO-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    var row = new Array(hdrs.length).fill('');
    row[_medInvColIdx(hdrs, 'ID')] = id;
    row[_medInvColIdx(hdrs, 'ProductoIngresos')] = producto;
    row[_medInvColIdx(hdrs, 'SKU')] = sku;
    row[_medInvColIdx(hdrs, 'NombreMedicamento')] = full.descripcion;
    row[_medInvColIdx(hdrs, 'CantidadPorUnidad')] = cantidad;
    row[_medInvColIdx(hdrs, 'Activo')] = 'TRUE';
    sh.appendRow(row);
    try { logAudit(body.usuario || '', 'Combos', 'Alta', id, '', '', producto + ' -> ' + full.descripcion); } catch (e) {}
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

/* ══════════════════════════════════════════════════════════════
   ÓRDENES DE COMPRA — liga Inventario ↔ Cuentas por Pagar
   ------------------------------------------------------------
   Reemplaza la vieja "Compras" (entrada rápida que solo tocaba
   inventario, sin generar ninguna cuenta por pagar). Una Orden de
   Compra SIEMPRE crea su CxP (vía saveCxP, sin modificar ese sistema)
   y por separado decide si el stock entra YA (recibida al momento)
   o se queda pendiente hasta que alguien la marque como recibida —
   recepción y pago son dos interruptores independientes, en
   cualquier orden, igual que en un ERP real:
     - Recibida ahora, pagar después: se registra la Entrada de
       inventario de inmediato y la CxP queda pendiente de pago.
     - Pagar antes de recibir: la CxP se puede pagar normal (con
       todo lo que ya existe: bancos, conciliación, créditos de
       proveedor) y el stock se queda en 0 hasta "Marcar recibida".
   ══════════════════════════════════════════════════════════════ */
var MEDINV_OC_TAB = 'Ordenes_Compra';
var MEDINV_OC_LINEAS_TAB = 'Ordenes_Compra_Lineas';
var MEDINV_OC_HEADERS = ['ID', 'Fecha', 'Proveedor', 'EstadoRecepcion', 'FechaRecepcion', 'Total', 'CxPId', 'CxPRowNum', 'Usuario', 'Notas'];
var MEDINV_OC_LINEAS_HEADERS = ['OrdenID', 'SKU', 'Nombre', 'Cantidad', 'CostoUnitario', 'Subtotal'];

/* ── Setup (correr UNA VEZ desde el editor de Apps Script) ───────── */
function setupOrdenesCompra() {
  var ss = _medInvSS();
  var creadas = [];
  if (!ss.getSheetByName(MEDINV_OC_TAB)) {
    var sh = ss.insertSheet(MEDINV_OC_TAB);
    sh.getRange(1, 1, 1, MEDINV_OC_HEADERS.length).setValues([MEDINV_OC_HEADERS]);
    sh.getRange(1, 1, 1, MEDINV_OC_HEADERS.length).setFontWeight('bold').setBackground('#1a252f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    var estCol = MEDINV_OC_HEADERS.indexOf('EstadoRecepcion') + 1;
    sh.getRange(2, estCol, 2000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Pendiente de recibir', 'Recibida'], true).setAllowInvalid(true).build()
    );
    sh.autoResizeColumns(1, MEDINV_OC_HEADERS.length);
    creadas.push(MEDINV_OC_TAB);
  }
  if (!ss.getSheetByName(MEDINV_OC_LINEAS_TAB)) {
    var sh2 = ss.insertSheet(MEDINV_OC_LINEAS_TAB);
    sh2.getRange(1, 1, 1, MEDINV_OC_LINEAS_HEADERS.length).setValues([MEDINV_OC_LINEAS_HEADERS]);
    sh2.getRange(1, 1, 1, MEDINV_OC_LINEAS_HEADERS.length).setFontWeight('bold').setBackground('#1a252f').setFontColor('#ffffff');
    sh2.setFrozenRows(1);
    sh2.autoResizeColumns(1, MEDINV_OC_LINEAS_HEADERS.length);
    creadas.push(MEDINV_OC_LINEAS_TAB);
  }
  return { ok: true, creadas: creadas };
}

function _ocNextId(sh) {
  var lr = sh.getLastRow();
  if (lr < 2) return 'OC-00001';
  var last = String(sh.getRange(lr, 1).getValue() || '');
  var m = last.match(/OC-(\d+)/);
  return 'OC-' + String((m ? parseInt(m[1], 10) : 0) + 1).padStart(5, '0');
}

function crearOrdenCompra(body) {
  try {
    setupOrdenesCompra();
    var proveedor = String(body.proveedor || '').trim();
    var lineasIn = body.lineas || [];
    if (!proveedor) return { ok: false, error: 'El proveedor es obligatorio.' };
    if (!lineasIn.length) return { ok: false, error: 'Agrega al menos un producto.' };

    var lineasValidas = [];
    var total = 0;
    for (var i = 0; i < lineasIn.length; i++) {
      var l = lineasIn[i];
      var sku = String(l.sku || '').trim();
      var cantidad = Number(l.cantidad) || 0;
      var costo = Number(l.costoUnitario) || 0;
      if (!sku || cantidad <= 0) continue;
      var prod = _bdProdFindBySku(sku);
      if (!prod) continue;
      // Comprar un producto lo vuelve inventariable desde ya (aunque la
      // mercancía no se haya recibido) — así queda listo para el stock.
      var invColsOc = _bdProdEnsureInventarioCols(prod.sh);
      prod.sh.getRange(prod.rowNum, invColsOc.Inventariable).setValue(true);
      var subtotal = cantidad * costo;
      total += subtotal;
      lineasValidas.push({ sku: sku, nombre: prod.descripcion, cantidad: cantidad, costoUnitario: costo, subtotal: subtotal });
    }
    if (!lineasValidas.length) return { ok: false, error: 'Ninguna línea es válida — revisa que el SKU exista y la cantidad sea mayor a 0.' };

    var fecha = body.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var recibida = body.recibida === true || body.recibida === 'true';
    var vencimiento = body.vencimiento || '';

    // 1. La orden SIEMPRE genera su cuenta por pagar — recibir mercancía y
    // pagarla son cosas independientes, pero toda compra se debe.
    var conceptoCxP = lineasValidas.map(function (l) { return l.nombre + ' x' + l.cantidad; }).join(', ');
    var cxpRes = saveCxP({
      proveedor: proveedor, contable: 'Costo', tipo: 'Variable',
      subtipo: body.subtipo || 'Medicamentos e Insumos',
      concepto: 'Orden de compra (pendiente) — ' + conceptoCxP,
      monto: total, vencimiento: vencimiento,
      mes: (vencimiento || fecha).substring(0, 7),
      notas: body.notas || '', usuario: body.usuario || ''
    });
    if (!cxpRes.ok) return { ok: false, error: 'No se pudo crear la cuenta por pagar: ' + cxpRes.error };

    // 2. Guardar la orden y sus líneas
    var sh = _medInvSheet(MEDINV_OC_TAB);
    var ocId = _ocNextId(sh);
    var row = new Array(MEDINV_OC_HEADERS.length).fill('');
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'ID')] = ocId;
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'Fecha')] = fecha;
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'Proveedor')] = proveedor;
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'EstadoRecepcion')] = recibida ? 'Recibida' : 'Pendiente de recibir';
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'FechaRecepcion')] = recibida ? fecha : '';
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'Total')] = total;
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'CxPId')] = cxpRes.id;
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'CxPRowNum')] = cxpRes.rowNum;
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'Usuario')] = body.usuario || '';
    row[_medInvColIdx(MEDINV_OC_HEADERS, 'Notas')] = body.notas || '';
    sh.appendRow(row);

    var lineasSh = _medInvSheet(MEDINV_OC_LINEAS_TAB);
    var lineRows = lineasValidas.map(function (l) { return [ocId, l.sku, l.nombre, l.cantidad, l.costoUnitario, l.subtotal]; });
    lineasSh.getRange(lineasSh.getLastRow() + 1, 1, lineRows.length, lineRows[0].length).setValues(lineRows);

    // 3. Si ya se recibió la mercancía, mover el inventario de una vez
    var movimientos = [];
    if (recibida) {
      lineasValidas.forEach(function (l) {
        movimientos.push(_registrarMovimientoInventario({
          sku: l.sku, nombre: l.nombre, tipo: 'Entrada', cantidad: l.cantidad,
          motivo: 'Compra', referencia: ocId + ' — ' + proveedor, costoUnitario: l.costoUnitario,
          usuario: body.usuario || '', notas: 'Orden de compra ' + ocId
        }));
      });
    }

    try { logAudit(body.usuario || '', 'Inventario', 'OrdenCompra', ocId, '', '', proveedor + ' | $' + total.toFixed(2)); } catch (e) {}
    return { ok: true, id: ocId, total: total, cxpId: cxpRes.id, recibida: recibida, movimientos: movimientos };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Lee el estado Pagado (col N, índice 14 en base 1) directo de Egresos2026
// para no duplicar ese dato en Ordenes_Compra — una sola fuente de verdad.
function _ocLeerEstadoPagoCxP(rowNums) {
  var out = {};
  if (!rowNums.length) return out;
  try {
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var sh = ss.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
    if (!sh) return out;
    rowNums.forEach(function (rn) {
      if (!rn) return;
      try { out[rn] = sh.getRange(rn, 14).getValue() === true; } catch (e) {}
    });
  } catch (e) {}
  return out;
}

function readOrdenesCompra() {
  try {
    setupOrdenesCompra();
    var sh = _medInvSheet(MEDINV_OC_TAB);
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var rows = [];
    for (var i = data.length - 1; i >= 1; i--) { // más reciente primero
      var r = data[i];
      if (!String(r[0] || '').trim()) continue;
      rows.push({
        _rowNum: i + 1,
        id: String(r[_medInvColIdx(hdrs, 'ID')] || ''),
        fecha: String(r[_medInvColIdx(hdrs, 'Fecha')] || ''),
        proveedor: String(r[_medInvColIdx(hdrs, 'Proveedor')] || ''),
        estadoRecepcion: String(r[_medInvColIdx(hdrs, 'EstadoRecepcion')] || ''),
        fechaRecepcion: String(r[_medInvColIdx(hdrs, 'FechaRecepcion')] || ''),
        total: Number(r[_medInvColIdx(hdrs, 'Total')]) || 0,
        cxpId: String(r[_medInvColIdx(hdrs, 'CxPId')] || ''),
        cxpRowNum: Number(r[_medInvColIdx(hdrs, 'CxPRowNum')]) || 0,
        usuario: String(r[_medInvColIdx(hdrs, 'Usuario')] || ''),
        notas: String(r[_medInvColIdx(hdrs, 'Notas')] || '')
      });
      if (rows.length >= 300) break;
    }

    var pagadoMap = _ocLeerEstadoPagoCxP(rows.map(function (r) { return r.cxpRowNum; }));
    rows.forEach(function (r) { r.pagado = !!pagadoMap[r.cxpRowNum]; });

    var lineasSh = _medInvSheet(MEDINV_OC_LINEAS_TAB);
    var lineasData = lineasSh.getDataRange().getValues();
    var lineasPorOC = {};
    for (var li = 1; li < lineasData.length; li++) {
      var lr = lineasData[li];
      var ocid = String(lr[0] || '').trim();
      if (!ocid) continue;
      if (!lineasPorOC[ocid]) lineasPorOC[ocid] = [];
      lineasPorOC[ocid].push({ sku: String(lr[1] || ''), nombre: String(lr[2] || ''), cantidad: Number(lr[3]) || 0, costoUnitario: Number(lr[4]) || 0, subtotal: Number(lr[5]) || 0 });
    }
    rows.forEach(function (r) { r.lineas = lineasPorOC[r.id] || []; });

    return { ok: true, rows: rows };
  } catch (ex) { return { ok: false, error: ex.message, rows: [] }; }
}

function marcarOrdenRecibida(body) {
  try {
    var rowNum = Number(body.rowNum);
    if (!rowNum) return { ok: false, error: 'Falta rowNum' };
    var sh = _medInvSheet(MEDINV_OC_TAB);
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var data = sh.getRange(rowNum, 1, 1, hdrs.length).getValues()[0];
    var ocId = String(data[_medInvColIdx(hdrs, 'ID')] || '');
    var estadoActual = String(data[_medInvColIdx(hdrs, 'EstadoRecepcion')] || '');
    if (estadoActual === 'Recibida') return { ok: false, error: 'Esta orden ya está marcada como recibida.' };
    var proveedor = String(data[_medInvColIdx(hdrs, 'Proveedor')] || '');

    var lineasSh = _medInvSheet(MEDINV_OC_LINEAS_TAB);
    var lineasData = lineasSh.getDataRange().getValues();
    var movimientos = [];
    for (var i = 1; i < lineasData.length; i++) {
      var lr = lineasData[i];
      if (String(lr[0] || '').trim() !== ocId) continue;
      movimientos.push(_registrarMovimientoInventario({
        sku: String(lr[1] || ''), nombre: String(lr[2] || ''), tipo: 'Entrada', cantidad: Number(lr[3]) || 0,
        motivo: 'Compra', referencia: ocId + ' — ' + proveedor, costoUnitario: Number(lr[4]) || 0,
        usuario: body.usuario || '', notas: 'Recepción de orden de compra ' + ocId
      }));
    }
    if (!movimientos.length) return { ok: false, error: 'No se encontraron líneas para esta orden.' };

    var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sh.getRange(rowNum, _medInvColIdx(hdrs, 'EstadoRecepcion') + 1).setValue('Recibida');
    sh.getRange(rowNum, _medInvColIdx(hdrs, 'FechaRecepcion') + 1).setValue(fechaHoy);
    try { logAudit(body.usuario || '', 'Inventario', 'OrdenRecibida', ocId, '', '', ''); } catch (e) {}
    return { ok: true, movimientos: movimientos };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Corrige el menú (hoja "Menu" en SHEET_ID) ────────────────────
   Medicamentos deja de ser un grupo con submenú — ahora es una vista
   directa (igual que Tratamientos/Estudios: filtra Catálogo General
   por Categoria=Medicamento vía el mismo redirect de navigateTo() que
   ya existía para 'prod-med...'). Compras/Movimientos/Combos se mudan
   a un grupo nuevo "Inventario" bajo Captura, porque ya operan sobre
   cualquier producto Inventariable, no solo medicamentos — vivir bajo
   "Medicamentos" ya no tenía sentido. La vieja "Compras" (entrada
   rápida sin CxP) se reemplaza por "Órdenes de Compra". Correr UNA
   VEZ desde el editor de Apps Script. */
function configurarMenuInventario() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Menu');
  if (!sh) return { ok: false, error: 'No se encontró la hoja Menu' };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var idxCol = hdrs.indexOf('ID'), padreCol = hdrs.indexOf('Padre'), tipoCol = hdrs.indexOf('Tipo'),
      fuenteCol = hdrs.indexOf('Fuente'), ordenCol = hdrs.indexOf('Orden');

  var filasABorrar = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][idxCol] || '').trim();
    if (id === 'meds-lista' || id === 'meds-compras') { filasABorrar.push(i + 1); continue; }
    if (id === 'prod-meds') {
      sh.getRange(i + 1, tipoCol + 1).setValue('vista');
      sh.getRange(i + 1, fuenteCol + 1).setValue('');
    }
    if (id === 'meds-movimientos') {
      sh.getRange(i + 1, padreCol + 1).setValue('inventario');
      sh.getRange(i + 1, ordenCol + 1).setValue(2);
    }
    if (id === 'meds-combos') {
      sh.getRange(i + 1, padreCol + 1).setValue('inventario');
      sh.getRange(i + 1, ordenCol + 1).setValue(3);
    }
  }
  filasABorrar.sort(function (a, b) { return b - a; }); // de abajo hacia arriba
  filasABorrar.forEach(function (rowNum) { sh.deleteRow(rowNum); });

  var nuevasFilas = [
    ['inventario', 'captura', 'Inventario', '', 'archive', 7, 'grupo', '', 'TRUE'],
    ['oc-lista', 'inventario', 'Órdenes de Compra', '', 'shopping-cart', 1, 'vista', 'oc-lista', 'TRUE']
  ];
  sh.getRange(sh.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);

  return { ok: true, filasBorradas: filasABorrar.length, filasAgregadas: nuevasFilas.length };
}

/* ── (Histórico) Corrige el menú — versión anterior, ya no se usa ── */
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
        motivo: 'Venta', referencia: opId, usuario: usuario || '',
        notas: 'Producto: ' + l.producto
      });
      resultados.push({ sku: c.sku, cantidad: cantidadADescontar, ok: r.ok, saldoNuevo: r.saldoNuevo, error: r.error });
    });
  });
  return { ok: true, descontados: resultados.length, detalle: resultados };
}
