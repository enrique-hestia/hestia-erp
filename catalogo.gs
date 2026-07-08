// catalogo.gs — Unificación de Catálogo de Productos
// Hestia Fertility ERP
//
// PASO 1: Ejecuta unificarCatalogo() → genera hoja Mapeo_Productos (solo lectura)
// PASO 2: Revisa el mapeo, corrige amarillos y rojos
// PASO 3 en adelante: se implementan después de aprobar el mapeo

var CAT_ID  = '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY';
var MP_ID   = '1qIc9x0JJrb0e2qVxrGiZqHA3CD36r0Wh';
var CNT_ID  = '1zH98lB4tJ5LA9FTCRqE1DaiFgGAjUFOg';
var MAPEO_TAB = 'Mapeo_Productos';

// ============================================================
// PASO 1: Genera el reporte de mapeo — CERO cambios en ningún archivo
// ============================================================
function unificarCatalogo() {
  Logger.log('Leyendo Mercado Pago...');
  var mpSS  = SpreadsheetApp.openById(MP_ID);
  var mpSh  = mpSS.getSheets()[0];
  var mpAll = mpSh.getDataRange().getValues();

  // Estructura MP confirmada por diagnóstico:
  // Datos desde fila 7 (índice 6)
  // Col 0=MLM_ID  Col 2=Nombre  Col 3=Precio  Col 4=Unidad  Col 7=SKU  Col 10=Sección
  var mpMap  = {};
  var mpList = [];
  for (var i = 6; i < mpAll.length; i++) {
    var r      = mpAll[i];
    var nombre = String(r[2]  || '').trim();
    var sku    = String(r[7]  || '').trim();
    if (!nombre || !sku) continue;
    var obj = {
      sku:     sku,
      nombre:  nombre,
      precio:  r[3],
      unidad:  String(r[4]  || '').trim(),
      seccion: String(r[10] || '').trim()
    };
    mpMap[_n(nombre)] = obj;
    mpList.push({ norm: _n(nombre), obj: obj });
  }
  Logger.log('MP: ' + mpList.length + ' productos');

  Logger.log('Leyendo Cntadigital...');
  var cntSS  = SpreadsheetApp.openById(CNT_ID);
  var cntSh  = cntSS.getSheets()[0];
  var cntAll = cntSh.getDataRange().getValues();

  // Detectar columnas por nombre de encabezado
  var cntHdr = cntAll[0].map(function(h) { return String(h).trim().toUpperCase(); });
  function ci(name) { return cntHdr.indexOf(name); }

  var cntMap  = {};
  var cntList = [];
  for (var i = 1; i < cntAll.length; i++) {
    var r      = cntAll[i];
    var nombre = ci('PRODUCTO') >= 0 ? String(r[ci('PRODUCTO')] || '').trim() : '';
    if (!nombre) continue;
    var obj = {
      nombre:         nombre,
      unidad:         String(r[ci('UNIDAD')]          || ''),
      tipo:           String(r[ci('TIPO')]            || ''),
      ieps:           String(r[ci('IEPS')]            || ''),
      cuentaContable: String(r[ci('CUENTA_CONTABLE')] || ''),
      tasa:           String(r[ci('TASA')]            || ''),
      objetoImpuesto: String(r[ci('OBJETO_IMPUESTO')] || ''),
      claveSat:       String(r[ci('CLAVE_SAT')]       || '')
    };
    cntMap[_n(nombre)] = obj;
    cntList.push({ norm: _n(nombre), obj: obj });
  }
  Logger.log('CNT: ' + cntList.length + ' productos');

  Logger.log('Leyendo Catálogo General...');
  var catSS  = SpreadsheetApp.openById(CAT_ID);
  var catSh  = catSS.getSheets()[0];
  var catAll = catSh.getDataRange().getValues();
  // Cols: 0=ProductoID  1=SKU  2=Descripcion  3=Categoria  4=Tipo
  Logger.log('Catálogo: ' + (catAll.length - 1) + ' productos');

  // ── Construir mapeo ──────────────────────────────────────
  var HDR = [
    'SKU_NUEVO', 'SKU_ANTERIOR', 'DESCRIPCION_MP', 'DESCRIPCION_ANTERIOR',
    'CATEGORIA', 'TIPO_CAT', 'PRECIO_MP', 'UNIDAD_MP', 'SECCION_MP',
    'IEPS', 'CLAVE_SAT', 'UNIDAD_SAT', 'TIPO_SAT', 'TASA', 'OBJETO_IMPUESTO',
    'CUENTA_CONTABLE', 'MATCH_MP', 'MATCH_CNT', 'CONFIANZA', 'NOTAS'
  ];
  var rows = [HDR];
  var usedSkus        = {};
  var catMatchedNorms = {};

  // Pase 1: MP como base (225 productos)
  for (var m = 0; m < mpList.length; m++) {
    var mpNorm = mpList[m].norm;
    var mp     = mpList[m].obj;

    var catHit  = _bestMatch(mpNorm, catAll);
    var cntHit  = cntMap[mpNorm] || _fuzzy(mpNorm, cntList, 0.60);

    var catDesc  = catHit ? catHit.desc      : '';
    var catSku2  = catHit ? catHit.sku       : '';
    var catCat   = catHit ? catHit.cat       : '';
    var catTipo  = catHit ? catHit.tipo      : '';
    var matchMp  = catHit ? catHit.matchType : 'NINGUNO';
    var conf     = catHit ? catHit.confianza : 'SIN_MATCH';
    var cntType  = cntHit ? (cntMap[mpNorm] ? 'EXACTO' : 'PARCIAL') : 'NINGUNO';

    if (catHit) catMatchedNorms[_n(catHit.desc)] = true;
    usedSkus[mp.sku] = true;

    var nota = catHit ? '' : 'En MP pero sin match en Catálogo General';
    rows.push([
      mp.sku,
      catSku2,
      mp.nombre,
      catDesc,
      catCat,
      catTipo,
      mp.precio,
      mp.unidad,
      mp.seccion,
      cntHit ? cntHit.ieps           : '',
      cntHit ? cntHit.claveSat       : '',
      cntHit ? cntHit.unidad         : '',
      cntHit ? cntHit.tipo           : '',
      cntHit ? cntHit.tasa           : '',
      cntHit ? cntHit.objetoImpuesto : '',
      cntHit ? cntHit.cuentaContable : '',
      matchMp,
      cntType,
      conf,
      nota
    ]);
  }

  // Pase 2: productos en Catálogo General sin match en MP → SKU nuevo corto
  var counters = {};
  for (var i = 1; i < catAll.length; i++) {
    var r    = catAll[i];
    var desc = String(r[2] || '').trim();
    if (!desc) continue;
    if (catMatchedNorms[_n(desc)]) continue;

    var cat    = String(r[3] || '').trim();
    var tipo   = String(r[4] || '').trim();
    var prefix = _prefix(cat || tipo);
    if (!counters[prefix]) counters[prefix] = 1;
    var newSku = prefix + '-' + _pad(counters[prefix]++);
    while (usedSkus[newSku]) { newSku = prefix + '-' + _pad(counters[prefix]++); }
    usedSkus[newSku] = true;

    var cntHit = cntMap[_n(desc)] || _fuzzy(_n(desc), cntList, 0.60);

    rows.push([
      newSku,
      String(r[1] || ''),
      '',
      desc,
      cat,
      tipo,
      '',
      cntHit ? cntHit.unidad : '',
      '',
      cntHit ? cntHit.ieps           : '',
      cntHit ? cntHit.claveSat       : '',
      cntHit ? cntHit.unidad         : '',
      cntHit ? cntHit.tipo           : '',
      cntHit ? cntHit.tasa           : '',
      cntHit ? cntHit.objetoImpuesto : '',
      cntHit ? cntHit.cuentaContable : '',
      'NINGUNO',
      cntHit ? 'PARCIAL' : 'NINGUNO',
      'REVISAR',
      'Sin match en MP — asigna SKU de MP en SKU_NUEVO si existe, o deja el generado'
    ]);
  }

  // ── Escribir hoja Mapeo_Productos ───────────────────────
  var old = catSS.getSheetByName(MAPEO_TAB);
  if (old) catSS.deleteSheet(old);
  var outSh = catSS.insertSheet(MAPEO_TAB);

  outSh.getRange(1, 1, rows.length, HDR.length).setValues(rows);
  outSh.getRange(1, 1, 1, HDR.length)
       .setBackground('#1a252f').setFontColor('#ffffff').setFontWeight('bold');
  outSh.setFrozenRows(1);
  outSh.setFrozenColumns(3);

  // Colorear por confianza (col 19 = CONFIANZA, col 17 = MATCH_MP)
  for (var rx = 2; rx <= rows.length; rx++) {
    var conf2  = rows[rx - 1][18]; // CONFIANZA
    var matchP = rows[rx - 1][16]; // MATCH_MP
    var bg;
    if      (matchP === 'NINGUNO')  bg = '#f8d7da'; // rojo — sin match MP
    else if (conf2  === 'ALTA')     bg = '#d4edda'; // verde — automático
    else if (conf2  === 'MEDIA')    bg = '#fff3cd'; // amarillo — revisar
    else                            bg = '#ffeeba'; // ámbar — baja confianza
    outSh.getRange(rx, 1, 1, HDR.length).setBackground(bg);
  }
  outSh.autoResizeColumns(1, HDR.length);

  // ── Estadísticas ────────────────────────────────────────
  var nAlta    = 0, nMedia = 0, nSinMatch = 0;
  for (var rx = 1; rx < rows.length; rx++) {
    var c = rows[rx][18];
    var p = rows[rx][16];
    if      (p === 'NINGUNO') nSinMatch++;
    else if (c === 'ALTA')    nAlta++;
    else                      nMedia++;
  }

  Logger.log('========== REPORTE MAPEO ==========');
  Logger.log('Total filas: '                          + (rows.length - 1));
  Logger.log('VERDE  — ALTA confianza (auto):         ' + nAlta);
  Logger.log('AMARILLO — MEDIA confianza (revisar):   ' + nMedia);
  Logger.log('ROJO   — Sin match en MP (SKU nuevo):  ' + nSinMatch);
  Logger.log('Hoja creada: ' + MAPEO_TAB + ' en Catálogo General');
  Logger.log('Abre esa hoja, revisa, y avísame cuando estés listo para el siguiente paso.');
}

// ============================================================
// PASO 3: Aplica el mapeo aprobado al Catálogo General
// Requisito: haber revisado y corregido Mapeo_Productos
// NO toca BD_Ingresos
// ============================================================
function aplicarCatalogoUnificado() {
  var catSS  = SpreadsheetApp.openById(CAT_ID);
  var mapSh  = catSS.getSheetByName(MAPEO_TAB);
  if (!mapSh) throw new Error('Ejecuta primero unificarCatalogo()');

  var mapData = mapSh.getDataRange().getValues();
  var mHdr    = mapData[0];
  function mc(n) { return mHdr.indexOf(n); }

  // Asegurar columnas fiscales en Catálogo General
  var catSh   = catSS.getSheets()[0];
  var catHdrR = catSh.getRange(1, 1, 1, catSh.getLastColumn()).getValues()[0];
  var newCols = ['SKU_ANTERIOR','DESCRIPCION_MP','PRECIO_MP','UNIDAD_MP','SECCION_MP',
                 'IEPS','CLAVE_SAT','UNIDAD_SAT','TIPO_SAT','TASA','OBJETO_IMPUESTO','CUENTA_CONTABLE'];
  newCols.forEach(function(c) {
    if (catHdrR.indexOf(c) < 0) {
      catSh.getRange(1, catSh.getLastColumn() + 1).setValue(c);
      catHdrR.push(c);
    }
  });
  function cc(n) { return catHdrR.indexOf(n); }

  // Indexar catálogo actual por descripción y SKU
  var catData  = catSh.getDataRange().getValues();
  var idxDesc  = {}, idxSku = {};
  for (var i = 1; i < catData.length; i++) {
    var d = String(catData[i][2] || '').trim();
    var s = String(catData[i][1] || '').trim();
    if (d) idxDesc[_n(d)] = i + 1; // fila 1-based en sheet
    if (s) idxSku[s]      = i + 1;
  }

  var updated = 0, inactivos = 0;

  // Actualizar cada fila aprobada del mapeo
  for (var m = 1; m < mapData.length; m++) {
    var row     = mapData[m];
    var skuNvo  = String(row[mc('SKU_NUEVO')]            || '').trim();
    var skuAnt  = String(row[mc('SKU_ANTERIOR')]         || '').trim();
    var descMp  = String(row[mc('DESCRIPCION_MP')]       || '').trim();
    var descAnt = String(row[mc('DESCRIPCION_ANTERIOR')] || '').trim();
    if (!skuNvo || !descMp) continue;

    var shRow = idxDesc[_n(descAnt)] || idxDesc[_n(descMp)] || idxSku[skuAnt];

    if (!shRow) {
      // Producto en MP que NO existe en Catálogo General → agregar fila nueva
      if (!skuNvo || !descMp) continue;
      var newRow = new Array(catHdrR.length).fill('');
      newRow[cc('SKU')]            = skuNvo;
      newRow[cc('Descripcion')]    = descMp;
      newRow[cc('Categoria')]      = String(row[mc('SECCION_MP')] || '');
      newRow[cc('Activo')]         = 'TRUE';
      newRow[cc('FechaCreacion')]  = new Date();
      if (cc('DESCRIPCION_MP')    >= 0) newRow[cc('DESCRIPCION_MP')]    = descMp;
      if (cc('PRECIO_MP')         >= 0) newRow[cc('PRECIO_MP')]         = row[mc('PRECIO_MP')];
      if (cc('UNIDAD_MP')         >= 0) newRow[cc('UNIDAD_MP')]         = row[mc('UNIDAD_MP')];
      if (cc('SECCION_MP')        >= 0) newRow[cc('SECCION_MP')]        = row[mc('SECCION_MP')];
      if (cc('IEPS')              >= 0) newRow[cc('IEPS')]              = row[mc('IEPS')];
      if (cc('CLAVE_SAT')         >= 0) newRow[cc('CLAVE_SAT')]         = row[mc('CLAVE_SAT')];
      if (cc('UNIDAD_SAT')        >= 0) newRow[cc('UNIDAD_SAT')]        = row[mc('UNIDAD_SAT')];
      if (cc('TIPO_SAT')          >= 0) newRow[cc('TIPO_SAT')]          = row[mc('TIPO_SAT')];
      if (cc('TASA')              >= 0) newRow[cc('TASA')]              = row[mc('TASA')];
      if (cc('OBJETO_IMPUESTO')   >= 0) newRow[cc('OBJETO_IMPUESTO')]   = row[mc('OBJETO_IMPUESTO')];
      if (cc('CUENTA_CONTABLE')   >= 0) newRow[cc('CUENTA_CONTABLE')]   = row[mc('CUENTA_CONTABLE')];
      catSh.appendRow(newRow);
      updated++;
      continue;
    }

    // Producto existente → actualizar
    catSh.getRange(shRow, cc('SKU') + 1).setValue(skuNvo);
    catSh.getRange(shRow, cc('Descripcion') + 1).setValue(descMp);
    if (cc('SKU_ANTERIOR')   >= 0 && skuAnt)  catSh.getRange(shRow, cc('SKU_ANTERIOR')   + 1).setValue(skuAnt);
    if (cc('DESCRIPCION_MP') >= 0)             catSh.getRange(shRow, cc('DESCRIPCION_MP') + 1).setValue(descMp);
    var fiscales = ['PRECIO_MP','UNIDAD_MP','SECCION_MP','IEPS','CLAVE_SAT',
                    'UNIDAD_SAT','TIPO_SAT','TASA','OBJETO_IMPUESTO','CUENTA_CONTABLE'];
    fiscales.forEach(function(f) {
      var val = row[mc(f)];
      if (cc(f) >= 0 && val !== '' && val !== null && val !== undefined)
        catSh.getRange(shRow, cc(f) + 1).setValue(val);
    });
    updated++;
  }

  // Marcar como inactivos los productos sin SKU de MP
  // EXCEPCIÓN: si en Mapeo_Productos la columna NOTAS dice SANTANDER (u otro método
  // de pago externo a MP), el producto se conserva activo con su SKU generado
  var catData2    = catSh.getDataRange().getValues();
  var skusActivos = {}; // SKUs que deben permanecer activos
  for (var m = 1; m < mapData.length; m++) {
    var s    = String(mapData[m][mc('SKU_NUEVO')] || '').trim();
    var nota = String(mapData[m][mc('NOTAS')]     || '').trim().toUpperCase();
    if (!s) continue;
    // Activo si tiene match en MP O si la nota indica pago externo
    if (nota !== 'INACTIVO') skusActivos[s] = true;
  }
  for (var i = 1; i < catData2.length; i++) {
    var sku = String(catData2[i][cc('SKU')]    || '').trim();
    var act = String(catData2[i][cc('Activo')] || '').trim();
    if (!sku) continue;
    if (!skusActivos[sku] && act.toUpperCase() !== 'FALSE') {
      catSh.getRange(i + 1, cc('Activo') + 1).setValue('FALSE');
      inactivos++;
    }
  }

  Logger.log('========== CATÁLOGO ACTUALIZADO ==========');
  Logger.log('Actualizados con datos MP: ' + updated);
  Logger.log('Marcados inactivos (sin match MP): ' + inactivos);
  Logger.log('SIGUIENTE: Ejecuta generarReporteHistorial()');
}

// ============================================================
// PASO 4: Reporte de correcciones en BD_Ingresos (SOLO LECTURA)
// ============================================================
function generarReporteHistorial() {
  var catSS   = SpreadsheetApp.openById(CAT_ID);
  var mapSh   = catSS.getSheetByName(MAPEO_TAB);
  if (!mapSh) throw new Error('Ejecuta primero unificarCatalogo()');

  var mapData = mapSh.getDataRange().getValues();
  var mHdr    = mapData[0];
  function mc(n) { return mHdr.indexOf(n); }

  // Diccionario: norm(descripción anterior) → {antes, despues}
  var fixMap = {};
  for (var m = 1; m < mapData.length; m++) {
    var row     = mapData[m];
    var descMp  = String(row[mc('DESCRIPCION_MP')]       || '').trim();
    var descAnt = String(row[mc('DESCRIPCION_ANTERIOR')] || '').trim();
    if (!descMp || !descAnt) continue;
    if (_n(descMp) === _n(descAnt)) continue; // sin cambio
    fixMap[_n(descAnt)] = { antes: descAnt, despues: descMp };
  }

  var mainSS  = SpreadsheetApp.openById(MAIN_ID);
  var bdSh    = mainSS.getSheetByName(BD_TAB);
  if (!bdSh) throw new Error('No encontré hoja ' + BD_TAB);

  var bdData  = bdSh.getDataRange().getValues();
  var bdHdr   = bdData[0];
  var prodIdx = _findCol(bdHdr, ['Producto','PRODUCTO','producto']);
  if (prodIdx < 0) throw new Error('No encontré columna Producto en ' + BD_TAB);
  Logger.log('Columna Producto en BD_Ingresos: col ' + (prodIdx + 1));

  var report = [['FILA_BD','DESCRIPCION_ACTUAL','DESCRIPCION_NUEVA','APLICAR']];
  for (var i = 1; i < bdData.length; i++) {
    var prod = String(bdData[i][prodIdx] || '').trim();
    if (!prod) continue;
    var hit = fixMap[_n(prod)];
    if (!hit) continue;
    report.push([i + 1, prod, hit.despues, 'SI']);
  }

  var old = mainSS.getSheetByName(FIX_TAB);
  if (old) mainSS.deleteSheet(old);
  var fixSh = mainSS.insertSheet(FIX_TAB);
  fixSh.getRange(1, 1, report.length, 4).setValues(report);
  fixSh.getRange(1, 1, 1, 4).setBackground('#1a252f').setFontColor('#ffffff').setFontWeight('bold');
  fixSh.setFrozenRows(1);
  fixSh.getRange(2, 1, report.length - 1, 4).setBackground('#d4edda');
  fixSh.autoResizeColumns(1, 4);

  Logger.log('========== REPORTE HISTORIAL ==========');
  Logger.log('Filas BD_Ingresos a corregir: ' + (report.length - 1));
  Logger.log('Todas marcadas APLICAR=SI — cambia a NO las que no quieras modificar');
  Logger.log('SIGUIENTE: Revisa Fix_Historial y ejecuta aplicarFixHistorial()');
}

// ============================================================
// PASO 5: Aplica correcciones en BD_Ingresos con backup automático
// ============================================================
function aplicarFixHistorial() {
  var mainSS  = SpreadsheetApp.openById(MAIN_ID);
  var fixSh   = mainSS.getSheetByName(FIX_TAB);
  if (!fixSh) throw new Error('Ejecuta primero generarReporteHistorial()');

  var fixData = fixSh.getDataRange().getValues();
  var toFix   = [];
  for (var i = 1; i < fixData.length; i++) {
    if (String(fixData[i][3] || '').trim().toUpperCase() !== 'SI') continue;
    toFix.push({ sheetRow: Number(fixData[i][0]), antes: String(fixData[i][1] || '').trim(),
                 despues: String(fixData[i][2] || '').trim() });
  }
  if (!toFix.length) { Logger.log('No hay filas con APLICAR=SI.'); return; }

  // Backup automático
  var bdSh  = mainSS.getSheetByName(BD_TAB);
  var oldBk = mainSS.getSheetByName(BACKUP_TAB);
  if (oldBk) mainSS.deleteSheet(oldBk);
  bdSh.copyTo(mainSS).setName(BACKUP_TAB);
  Logger.log('Backup creado: ' + BACKUP_TAB);

  var bdHdr   = bdSh.getRange(1, 1, 1, bdSh.getLastColumn()).getValues()[0];
  var prodCol = _findCol(bdHdr, ['Producto','PRODUCTO','producto']) + 1;

  var ok = 0, skip = 0;
  toFix.forEach(function(f) {
    var cell   = bdSh.getRange(f.sheetRow, prodCol);
    var actual = String(cell.getValue() || '').trim();
    if (_n(actual) === _n(f.antes)) { cell.setValue(f.despues); ok++; }
    else { Logger.log('SKIP fila ' + f.sheetRow + ': esperaba "' + f.antes + '"'); skip++; }
  });

  Logger.log('========== FIX HISTORIAL ==========');
  Logger.log('Aplicados: ' + ok + ' | Omitidos: ' + skip);
  Logger.log('Si algo falló: copia ' + BACKUP_TAB + ' de vuelta a ' + BD_TAB);
}

// ============================================================
// PASO 6: Exportar catálogo unificado en formato Cntadigital
// Genera hoja "Exportar_Cntadigital" lista para descargar e importar
// ============================================================
function exportarParaCntadigital() {
  var catSS   = SpreadsheetApp.openById(CAT_ID);
  var mapSh   = catSS.getSheetByName(MAPEO_TAB);
  if (!mapSh) throw new Error('Ejecuta primero unificarCatalogo()');

  var mapData = mapSh.getDataRange().getValues();
  var mHdr    = mapData[0];
  function mc(n) { return mHdr.indexOf(n); }

  // Leer Cntadigital original para copiar valores por defecto
  var cntSS   = SpreadsheetApp.openById(CNT_ID);
  var cntSh   = cntSS.getSheets()[0];
  var cntAll  = cntSh.getDataRange().getValues();
  var cntHdr2 = cntAll[0].map(function(h) { return String(h).trim().toUpperCase(); });
  function ci(n) { return cntHdr2.indexOf(n); }

  // Construir lookup Cntadigital por nombre normalizado
  var cntIndex = {};
  for (var i = 1; i < cntAll.length; i++) {
    var nb = ci('PRODUCTO') >= 0 ? String(cntAll[i][ci('PRODUCTO')] || '').trim() : '';
    if (nb) cntIndex[_n(nb)] = cntAll[i];
  }

  // Encabezados en formato Cntadigital (30 columnas, igual que el original)
  var CNT_HDR = cntAll[0]; // usa exactamente los mismos encabezados del original

  var outRows = [CNT_HDR];
  var seq = 1;

  for (var m = 1; m < mapData.length; m++) {
    var row     = mapData[m];
    var skuNvo  = String(row[mc('SKU_NUEVO')]      || '').trim();
    var descMp  = String(row[mc('DESCRIPCION_MP')] || '').trim();
    var matchMp = String(row[mc('MATCH_MP')]       || '').trim();

    // Solo incluir productos con match en MP (los activos)
    if (!skuNvo || !descMp || matchMp === 'NINGUNO') continue;

    // Buscar fila original en Cntadigital para heredar valores fiscales
    var cntRow  = cntIndex[_n(descMp)] || cntIndex[_n(String(row[mc('DESCRIPCION_ANTERIOR')] || ''))];

    // Construir nueva fila en formato Cntadigital
    var newRow  = new Array(CNT_HDR.length).fill('');

    // Campos fijos o heredados
    newRow[ci('ID')]              = seq++;
    newRow[ci('CLAVE')]           = skuNvo;
    newRow[ci('CODIGO')]          = skuNvo;
    newRow[ci('PRODUCTO')]        = descMp;
    newRow[ci('PRECIO')]          = row[mc('PRECIO_MP')] || (cntRow ? cntRow[ci('PRECIO')] : '');
    newRow[ci('PRECIO_DOS')]      = cntRow ? cntRow[ci('PRECIO_DOS')]      : '0.0000';
    newRow[ci('DESCUENTO')]       = cntRow ? cntRow[ci('DESCUENTO')]       : '0.0000';
    newRow[ci('UNIDAD')]          = row[mc('UNIDAD_SAT')] || (cntRow ? cntRow[ci('UNIDAD')]  : 'SERVICIO');
    newRow[ci('TIPO')]            = row[mc('TIPO_SAT')]   || (cntRow ? cntRow[ci('TIPO')]    : 'Servicio');
    newRow[ci('DEPARTAMENTO')]    = cntRow ? cntRow[ci('DEPARTAMENTO')]    : '';
    newRow[ci('FAMILIA')]         = row[mc('SECCION_MP')] || (cntRow ? cntRow[ci('FAMILIA')] : '');
    newRow[ci('IEPS')]            = row[mc('IEPS')]        || (cntRow ? cntRow[ci('IEPS')]   : '0');
    newRow[ci('CUOTA_FIJA_IEPS')] = cntRow ? cntRow[ci('CUOTA_FIJA_IEPS')] : 'NO';
    newRow[ci('CUENTA_CONTABLE')] = row[mc('CUENTA_CONTABLE')] || (cntRow ? cntRow[ci('CUENTA_CONTABLE')] : '');
    newRow[ci('TASA')]            = row[mc('TASA')]        || (cntRow ? cntRow[ci('TASA')]   : 'Exento');
    newRow[ci('OBJETO_IMPUESTO')] = row[mc('OBJETO_IMPUESTO')] || (cntRow ? cntRow[ci('OBJETO_IMPUESTO')] : '02');
    newRow[ci('CLAVE_SAT')]       = row[mc('CLAVE_SAT')]   || (cntRow ? cntRow[ci('CLAVE_SAT')]  : '85121601');
    newRow[ci('NO_FACTURABLE')]   = 'NO';
    newRow[ci('ESTATUS')]         = 'SI';

    // Campos opcionales heredados
    var heredar = ['CLASIFICACION_EXTRA','CLASIFICACION_IEPS','CANTIDAD_MINIMA','CANTIDAD_MAXIMA',
                   'FRACCION_ARANCELARIA','UNIDAD_ADUANA','MARCA','SEGMENTO','PARTIDA','TIPO_PEDIDO'];
    heredar.forEach(function(f) {
      if (ci(f) >= 0 && cntRow) newRow[ci(f)] = cntRow[ci(f)] || '';
    });

    outRows.push(newRow);
  }

  // Escribir en hoja del spreadsheet de Cntadigital (reemplaza o crea)
  var expTab = 'Exportar_Unificado';
  var old    = cntSS.getSheetByName(expTab);
  if (old) cntSS.deleteSheet(old);
  var expSh  = cntSS.insertSheet(expTab);
  expSh.getRange(1, 1, outRows.length, CNT_HDR.length).setValues(outRows);
  expSh.getRange(1, 1, 1, CNT_HDR.length)
       .setBackground('#1a252f').setFontColor('#ffffff').setFontWeight('bold');
  expSh.setFrozenRows(1);
  expSh.autoResizeColumns(1, CNT_HDR.length);

  Logger.log('========== EXPORTAR CNTADIGITAL ==========');
  Logger.log('Productos exportados: ' + (outRows.length - 1));
  Logger.log('Hoja creada: "' + expTab + '" en el archivo de Cntadigital');
  Logger.log('Abre ese archivo, descarga como Excel (.xlsx) e importa en tu sistema de facturación');
}

// ============================================================
// HELPERS
// ============================================================

var MAIN_ID    = '1x_TE_YxLOwnBXKV_lA3Ss_EOSu1p61uTdUmw2zh_6uc';
var BD_TAB     = 'BD_Ingresos';
var FIX_TAB    = 'Fix_Historial';
var BACKUP_TAB = 'BD_Ingresos_BACKUP';

function _findCol(headers, aliases) {
  for (var k = 0; k < aliases.length; k++) {
    var idx = headers.indexOf(aliases[k]);
    if (idx >= 0) return idx;
  }
  return -1;
}

// ============================================================
// HELPERS (normalization and matching)
// ============================================================

function _n(s) {
  return String(s).toLowerCase()
    .replace(/[áàäâã]/g, 'a').replace(/[éèëê]/g, 'e')
    .replace(/[íìïî]/g, 'i').replace(/[óòöôõ]/g, 'o')
    .replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n')
    .replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function _prefix(s) {
  var clean = _n(s).replace(/ /g, '').toUpperCase();
  return (clean + 'XXX').substring(0, 3);
}

function _pad(n) {
  return String(n).padStart(3, '0');
}

function _tokens(norm) {
  return norm.split(' ').filter(function(t) { return t.length > 2; });
}

function _sim(a, b) {
  var ta = _tokens(a);
  var tb = _tokens(b);
  if (!ta.length || !tb.length) return 0;
  var common = ta.filter(function(t) { return tb.indexOf(t) >= 0; }).length;
  return (2 * common) / (ta.length + tb.length);
}

// Buscar mejor match en Catálogo General (catAll: col 2=desc, 1=sku, 3=cat, 4=tipo)
function _bestMatch(norm, catAll) {
  var best = null;
  var bestScore = 0;
  for (var i = 1; i < catAll.length; i++) {
    var desc = String(catAll[i][2] || '').trim();
    if (!desc) continue;
    var n = _n(desc);
    if (n === norm) {
      return {
        desc: desc,
        sku:  String(catAll[i][1] || ''),
        cat:  String(catAll[i][3] || ''),
        tipo: String(catAll[i][4] || ''),
        matchType: 'EXACTO',
        confianza: 'ALTA'
      };
    }
    var score = _sim(norm, n);
    if (score > bestScore && score >= 0.45) {
      bestScore = score;
      best = {
        desc: desc,
        sku:  String(catAll[i][1] || ''),
        cat:  String(catAll[i][3] || ''),
        tipo: String(catAll[i][4] || ''),
        matchType: 'PARCIAL',
        confianza: score >= 0.70 ? 'ALTA' : 'MEDIA'
      };
    }
  }
  return best;
}

// Buscar en lista con similitud mínima
function _fuzzy(norm, list, minScore) {
  var best = null;
  var bestScore = 0;
  for (var k = 0; k < list.length; k++) {
    var s = _sim(norm, list[k].norm);
    if (s > bestScore && s >= minScore) {
      bestScore = s;
      best = list[k].obj;
    }
  }
  return best;
}
