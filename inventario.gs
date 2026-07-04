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

/* ── Catálogo (filtro de BD_Productos: solo lo Inventariable) ─────── */
function readCatalogoMedicamentos() {
  try {
    var data = readProductos();
    if (!data.ok) return data;
    var rows = data.todosProductos.filter(function (p) { return p.inventariable; }).map(function (p) {
      return {
        productoId: p.id, sku: p.sku, nombre: p.descripcion, categoria: p.categoria,
        tipo: p.tipo, unidad: p.unidad, stockMinimo: p.stockMinimo, stockMaximo: p.stockMaximo,
        costoUnitario: p.costoUnitario, proveedorPreferido: p.proveedorPreferido,
        stockActual: p.stockActual, activo: p.activo, notas: p.notas
      };
    });
    rows.sort(function (a, b) { return a.nombre.toLowerCase() < b.nombre.toLowerCase() ? -1 : 1; });
    var kpis = {
      total: rows.length,
      bajoMinimo: rows.filter(function (r) { return r.activo && r.stockActual < r.stockMinimo; }).length,
      valorInventario: rows.reduce(function (s, r) { return s + r.stockActual * r.costoUnitario; }, 0)
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

/* ── Compras (entradas de mercancía) ─────────────────────────────── */
function registrarCompraMedicamento(body) {
  try {
    var sku = String(body.sku || '').trim();
    var cantidad = Number(body.cantidad) || 0;
    if (!sku || cantidad <= 0) return { ok: false, error: 'SKU y cantidad son obligatorios' };

    var prod = _bdProdFindBySku(sku);
    if (!prod) return { ok: false, error: 'SKU no encontrado en el catálogo' };

    var costoUnitario = Number(body.costoUnitario) || 0;
    var referencia = [body.proveedor, body.factura].filter(Boolean).join(' — ');
    var resultado = _registrarMovimientoInventario({
      sku: sku, nombre: prod.descripcion, tipo: 'Entrada', cantidad: cantidad,
      motivo: 'Compra', referencia: referencia, costoUnitario: costoUnitario,
      usuario: body.usuario || '', modulo: prod.categoria, notas: body.notas || ''
    });
    if (!resultado.ok) return resultado;

    // Actualiza el costo unitario vigente en el catálogo con el de la compra más reciente
    if (costoUnitario > 0) {
      var cols = _bdProdEnsureInventarioCols(prod.sh);
      prod.sh.getRange(prod.rowNum, cols.CostoUnitario).setValue(costoUnitario);
    }
    try { logAudit(body.usuario || '', 'Inventario', 'Compra', sku, 'cantidad', '', String(cantidad)); } catch (e) {}
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
    if (!full.inventariable) return { ok: false, error: 'Ese producto no está marcado como Inventariable — actívalo en Catálogo General antes de usarlo en un combo.' };

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
        motivo: 'Venta', referencia: opId, usuario: usuario || '',
        notas: 'Producto: ' + l.producto
      });
      resultados.push({ sku: c.sku, cantidad: cantidadADescontar, ok: r.ok, saldoNuevo: r.saldoNuevo, error: r.error });
    });
  });
  return { ok: true, descontados: resultados.length, detalle: resultados };
}
