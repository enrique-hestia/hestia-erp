// revisar_catalogo.gs — Revisión y corrección del catálogo de productos
//
// INSTRUCCIONES DE USO:
//   PASO 1: Ejecuta revisarCatalogoProductos()
//           → Genera hoja "Revision_Catalogo" con análisis completo (solo lectura)
//           → Revisa el reporte; edita la columna ACCION si quieres cambiar algo
//             (por ejemplo, cambiar INACTIVAR→OK para preservar un producto)
//   PASO 2: Ejecuta aplicarCorreccionesCatalogo()
//           → Aplica solo las filas con ACCION = ACTUALIZAR / REACTIVAR... / AGREGAR_FALTANTE
//           → Las filas INACTIVAR NO se aplican automáticamente — debes cambiar la
//             ACCION a INACTIVAR_CONFIRMAR en la hoja si quieres inactivar ese producto
//           → Hace backup automático en BD_Productos_BACKUP antes de cualquier cambio
//
// Reglas del catálogo vigente:
//  • Fuente de verdad = Mercado Pago (SKU y nombre de MP prevalecen)
//  • Productos en listas REPROVIDA o GrupoMedico en BD_Precios → ACTIVOS aunque no estén en MP
//  • Productos con "REPROVIDA" o "GRUPOMED" en campo Notas o Tipo → ACTIVOS
//  • Cualquier otro producto sin match en MP → se reporta como INACTIVAR (no se aplica auto)

var RC_MP_ID  = '1qIc9x0JJrb0e2qVxrGiZqHA3CD36r0Wh';
var RC_CAT_ID = '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY';
var RC_TAB    = 'Revision_Catalogo';
var RC_BACKUP = 'BD_Productos_BACKUP';
// Textos que identifican listas especiales
var RC_LISTAS = ['REPROVIDA','GRUPOMED','GRUPOMEDICO','SURROGACY','EXTERNO','EXTERNOS'];

// ─────────────────────────────────────────────────────────────────────────────
// PASO 1 — Análisis (no modifica nada)
// ─────────────────────────────────────────────────────────────────────────────
function revisarCatalogoProductos() {
  Logger.log('=== REVISIÓN CATÁLOGO — INICIO ===');

  // ── 1. Leer Mercado Pago ─────────────────────────────────────────────────
  var mpSS  = SpreadsheetApp.openById(RC_MP_ID);
  var mpSh  = mpSS.getSheets()[0];
  var mpRaw = mpSh.getDataRange().getValues();
  // Datos desde fila 7 (índice 6). Cols: 0=ID_MP, 2=Nombre, 3=Precio, 7=SKU, 10=Seccion
  var mpList   = [];
  var mpBySku  = {};
  var mpByNorm = {};
  for (var i = 6; i < mpRaw.length; i++) {
    var r = mpRaw[i];
    var nombre = String(r[2] || '').trim();
    if (!nombre) continue;
    var sku = String(r[7] || '').trim();
    var obj = {
      mlmId:   String(r[0] || '').trim(),
      sku:     sku,
      nombre:  nombre,
      precio:  parseFloat(r[3]) || 0,
      seccion: String(r[10] || '').trim()
    };
    mpList.push(obj);
    if (sku)  mpBySku[sku.toUpperCase()] = obj;
    mpByNorm[_rcN(nombre)] = obj;
  }
  Logger.log('MP: ' + mpList.length + ' productos cargados');

  // ── 2. Leer BD_Productos ─────────────────────────────────────────────────
  var catSS  = SpreadsheetApp.openById(RC_CAT_ID);
  var prodSh = catSS.getSheetByName('BD_Productos');
  var precSh = catSS.getSheetByName('BD_Precios');
  if (!prodSh) { Logger.log('ERROR: BD_Productos no encontrada en ' + RC_CAT_ID); return; }

  var prodData = prodSh.getDataRange().getValues();
  var prodHdr  = prodData[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
  function pc(name, fallback) {
    var idx = prodHdr.indexOf(name.toLowerCase());
    return idx >= 0 ? idx : fallback;
  }
  var iID   = pc('productoid', 0);
  var iSKU  = pc('sku',  1);
  var iDESC = pc('descripcion', 2);
  var iCAT  = pc('categoria', 3);
  var iTIPO = pc('tipo', 4);
  var iNOTA = pc('notas', 5);
  var iACT  = pc('activo', 6);

  // ── 3. Identificar productos con listas especiales ───────────────────────
  // A) Desde BD_Precios (columna Lista = índice 6)
  var enListaEspecial = {};
  if (precSh) {
    var precData = precSh.getDataRange().getValues();
    for (var pi = 1; pi < precData.length; pi++) {
      var pid   = String(precData[pi][0] || '').trim();
      var lista = String(precData[pi].length > 6 ? precData[pi][6] || '' : '').trim().toUpperCase();
      if (!pid) continue;
      for (var li = 0; li < RC_LISTAS.length; li++) {
        if (lista.indexOf(RC_LISTAS[li]) >= 0) {
          enListaEspecial[pid] = lista;
          break;
        }
      }
    }
  }
  Logger.log('Productos con lista especial en BD_Precios: ' + Object.keys(enListaEspecial).length);

  // B) Detectar por Notas o Tipo en BD_Productos
  function _eEspByNota(notas, tipo) {
    var combined = (String(notas||'') + ' ' + String(tipo||'')).toUpperCase();
    for (var li = 0; li < RC_LISTAS.length; li++) {
      if (combined.indexOf(RC_LISTAS[li]) >= 0) return RC_LISTAS[li];
    }
    return '';
  }

  // ── 4. Generar reporte ───────────────────────────────────────────────────
  var HDR = [
    'ACCION',
    'ProductoID', 'SKU_ACTUAL', 'DESCRIPCION_ACTUAL', 'CATEGORIA_ACTUAL', 'ACTIVO_ACTUAL',
    'SKU_MP', 'NOMBRE_MP', 'PRECIO_MP', 'SECCION_MP',
    'MATCH_TIPO', 'LISTA_ESPECIAL', 'NOTAS_REVISION'
  ];
  var rows    = [HDR];
  var mpUsado = {};

  for (var ci = 1; ci < prodData.length; ci++) {
    var r     = prodData[ci];
    var pid   = String(r[iID]  || '').trim();
    var sku   = String(r[iSKU] || '').trim();
    var desc  = String(r[iDESC]|| '').trim();
    var cat   = String(r[iCAT] || '').trim();
    var notas = String(r[iNOTA]|| '').trim();
    var tipo  = String(r[iTIPO]|| '').trim();
    var actv  = !(r[iACT] === false || String(r[iACT]).toUpperCase() === 'FALSE');
    if (!pid && !desc) continue;

    var mpHit = null;
    if (sku)  mpHit = mpBySku[sku.toUpperCase()];
    if (!mpHit && desc) mpHit = mpByNorm[_rcN(desc)];
    if (!mpHit && desc) mpHit = _rcFuzzy(desc, mpList, 0.75);

    var eEsp = enListaEspecial[pid] || _eEspByNota(notas, tipo) || '';
    var accion, matchTipo, notasRev;

    if (mpHit) {
      var mpKey = mpHit.sku ? mpHit.sku.toUpperCase() : _rcN(mpHit.nombre);
      mpUsado[mpKey] = true;
      if (mpHit.sku) mpUsado[_rcN(mpHit.nombre)] = true;

      matchTipo = (sku && sku.toUpperCase() === (mpHit.sku||'').toUpperCase()) ? 'SKU_EXACTO'
                : (_rcN(desc) === _rcN(mpHit.nombre))                           ? 'NOMBRE_EXACTO'
                                                                                 : 'NOMBRE_PARCIAL';
      var cambios = [];
      if (!actv)  cambios.push('Reactivar');
      if (mpHit.sku && sku !== mpHit.sku) cambios.push('SKU: "'+sku+'"→"'+mpHit.sku+'"');
      // Actualizar nombre solo si match es por SKU exacto
      if (matchTipo === 'SKU_EXACTO' && _rcN(desc) !== _rcN(mpHit.nombre)) {
        cambios.push('Nombre: "'+desc+'"→"'+mpHit.nombre+'"');
      }

      accion   = cambios.length ? 'ACTUALIZAR' : 'OK';
      notasRev = cambios.join(' | ');

    } else if (eEsp) {
      matchTipo = 'SIN_MATCH_MP';
      accion    = actv ? 'OK_LISTA_ESPECIAL' : 'REACTIVAR_LISTA_ESPECIAL';
      notasRev  = 'Lista: ' + eEsp + ' — permanece activo sin match MP';
      mpHit     = {};

    } else {
      matchTipo = 'SIN_MATCH_MP';
      accion    = actv ? 'INACTIVAR' : 'OK_INACTIVO';
      notasRev  = actv
        ? 'Sin match MP ni lista especial — REVISA si procede inactivar'
        : 'Ya inactivo, sin match MP';
      mpHit = {};
    }

    rows.push([
      accion, pid, sku, desc, cat, actv ? 'SI' : 'NO',
      mpHit.sku||'', mpHit.nombre||'', mpHit.precio||'', mpHit.seccion||'',
      matchTipo, eEsp, notasRev
    ]);
  }

  // ── 5. Productos en MP que no están en catálogo ──────────────────────────
  for (var mi = 0; mi < mpList.length; mi++) {
    var mp  = mpList[mi];
    var key = mp.sku ? mp.sku.toUpperCase() : _rcN(mp.nombre);
    if (mpUsado[key] || mpUsado[_rcN(mp.nombre)]) continue;

    rows.push([
      'AGREGAR_FALTANTE',
      '', '', '', '', 'NO',
      mp.sku, mp.nombre, mp.precio, mp.seccion,
      'EN_MP_NO_EN_CAT', '', 'Producto de MP sin entrada en BD_Productos'
    ]);
  }

  // ── 6. Escribir hoja de reporte ──────────────────────────────────────────
  var old = catSS.getSheetByName(RC_TAB);
  if (old) catSS.deleteSheet(old);
  var outSh = catSS.insertSheet(RC_TAB);
  outSh.getRange(1, 1, rows.length, HDR.length).setValues(rows);
  outSh.getRange(1, 1, 1, HDR.length)
    .setBackground('#1a252f').setFontColor('#ffffff').setFontWeight('bold');
  outSh.setFrozenRows(1);
  outSh.setFrozenColumns(1);

  var colores = {
    'OK':                      '#d4edda',
    'OK_INACTIVO':             '#e9ecef',
    'OK_LISTA_ESPECIAL':       '#cce5ff',
    'ACTUALIZAR':              '#fff3cd',
    'REACTIVAR_LISTA_ESPECIAL':'#d1ecf1',
    'INACTIVAR':               '#f8d7da',
    'AGREGAR_FALTANTE':        '#ffeeba'
  };
  for (var rx = 2; rx <= rows.length; rx++) {
    var a = rows[rx-1][0];
    outSh.getRange(rx, 1, 1, HDR.length).setBackground(colores[a] || '#ffffff');
  }
  outSh.autoResizeColumns(1, HDR.length);

  var noteCell = outSh.getRange(rows.length + 2, 1);
  noteCell.setValue(
    'INSTRUCCIONES: Filas INACTIVAR (rojo) NO se aplican automáticamente. ' +
    'Cámbialas a INACTIVAR_CONFIRMAR si deseas inactivarlas, o a OK para preservarlas. ' +
    'Luego ejecuta aplicarCorreccionesCatalogo().'
  );
  noteCell.setBackground('#fff9c4').setFontStyle('italic').setWrap(true);
  outSh.setColumnWidth(1, 220);

  var stats = {};
  rows.slice(1).forEach(function(r){ stats[r[0]] = (stats[r[0]]||0)+1; });
  Logger.log('\n=== RESUMEN ===');
  Object.keys(stats).sort().forEach(function(k){ Logger.log(k + ': ' + stats[k]); });
  Logger.log('\nHoja "' + RC_TAB + '" generada. Revisa y ejecuta aplicarCorreccionesCatalogo().');
}

// ─────────────────────────────────────────────────────────────────────────────
// PASO 2 — Aplicar correcciones
// ─────────────────────────────────────────────────────────────────────────────
function aplicarCorreccionesCatalogo() {
  var catSS = SpreadsheetApp.openById(RC_CAT_ID);
  var revSh = catSS.getSheetByName(RC_TAB);
  if (!revSh) {
    Logger.log('ERROR: Ejecuta primero revisarCatalogoProductos()');
    return;
  }

  var prodSh = catSS.getSheetByName('BD_Productos');
  if (!prodSh) { Logger.log('ERROR: BD_Productos no encontrada'); return; }

  // Backup automático
  var oldBk = catSS.getSheetByName(RC_BACKUP);
  if (oldBk) catSS.deleteSheet(oldBk);
  prodSh.copyTo(catSS).setName(RC_BACKUP);
  Logger.log('Backup creado: ' + RC_BACKUP);

  var revData = revSh.getDataRange().getValues();
  var rHdr    = revData[0].map(function(h){ return String(h||'').trim().toUpperCase(); });
  function rIdx(n){ return rHdr.indexOf(n.toUpperCase()); }

  var prodData = prodSh.getDataRange().getValues();
  var prodHdr  = prodData[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
  function pc(name, fallback) {
    var idx = prodHdr.indexOf(name.toLowerCase());
    return idx >= 0 ? idx : fallback;
  }
  var iID   = pc('productoid', 0);
  var iSKU  = pc('sku',  1);
  var iDESC = pc('descripcion', 2);
  var iCAT  = pc('categoria', 3);
  var iACT  = pc('activo', 6);

  var idxById   = {};
  var idxByNorm = {};
  var idxBySku  = {};
  for (var i = 1; i < prodData.length; i++) {
    var pid  = String(prodData[i][iID]  || '').trim();
    var desc = String(prodData[i][iDESC]|| '').trim();
    var sku  = String(prodData[i][iSKU] || '').trim().toUpperCase();
    if (pid)  idxById[pid]          = i + 1;
    if (desc) idxByNorm[_rcN(desc)] = i + 1;
    if (sku)  idxBySku[sku]         = i + 1;
  }

  var stats = {actualizar:0, reactivar:0, inactivar:0, agregar:0, skip:0, error:0};
  var log   = [];

  for (var ri = 1; ri < revData.length; ri++) {
    var row    = revData[ri];
    var accion = String(row[rIdx('ACCION')] || '').trim().toUpperCase();
    var doIt   = (accion === 'ACTUALIZAR'           ||
                  accion === 'REACTIVAR_LISTA_ESPECIAL' ||
                  accion === 'INACTIVAR_CONFIRMAR'  ||
                  accion === 'AGREGAR_FALTANTE');
    if (!doIt) { stats.skip++; continue; }

    var pid     = String(row[rIdx('PRODUCTOID')]         || '').trim();
    var descAct = String(row[rIdx('DESCRIPCION_ACTUAL')] || '').trim();
    var skuAct  = String(row[rIdx('SKU_ACTUAL')]         || '').trim();
    var skuMp   = String(row[rIdx('SKU_MP')]             || '').trim();
    var nomMp   = String(row[rIdx('NOMBRE_MP')]          || '').trim();
    var secMp   = String(row[rIdx('SECCION_MP')]         || '').trim();
    var notasRev= String(row[rIdx('NOTAS_REVISION')]     || '').trim();

    var shRow = (pid    ? idxById[pid]                   : null)
             || (descAct? idxByNorm[_rcN(descAct)]       : null)
             || (skuAct ? idxBySku[skuAct.toUpperCase()] : null)
             || (skuMp  ? idxBySku[skuMp.toUpperCase()]  : null);

    if (accion === 'ACTUALIZAR') {
      if (!shRow) { log.push('ERROR: no encontrada "' + (descAct||skuMp) + '"'); stats.error++; continue; }
      if (skuMp && iSKU >= 0)  prodSh.getRange(shRow, iSKU  + 1).setValue(skuMp);
      if (nomMp && notasRev.indexOf('Nombre:') >= 0 && iDESC >= 0) {
        prodSh.getRange(shRow, iDESC + 1).setValue(nomMp);
      }
      if (secMp && iCAT >= 0) {
        var catActual = String(prodData[shRow-1][iCAT]||'').trim();
        if (!catActual) prodSh.getRange(shRow, iCAT + 1).setValue(secMp);
      }
      if (iACT >= 0) prodSh.getRange(shRow, iACT + 1).setValue(true);
      stats.actualizar++;
      log.push('ACTUALIZADO: ' + (descAct||nomMp));

    } else if (accion === 'REACTIVAR_LISTA_ESPECIAL') {
      if (!shRow) { log.push('ERROR REACTIVAR: "' + descAct + '"'); stats.error++; continue; }
      if (iACT >= 0) prodSh.getRange(shRow, iACT + 1).setValue(true);
      stats.reactivar++;
      log.push('REACTIVADO: ' + descAct);

    } else if (accion === 'INACTIVAR_CONFIRMAR') {
      if (!shRow) { stats.skip++; continue; }
      if (iACT >= 0) prodSh.getRange(shRow, iACT + 1).setValue(false);
      stats.inactivar++;
      log.push('INACTIVADO: ' + descAct);

    } else if (accion === 'AGREGAR_FALTANTE') {
      if (!nomMp) { stats.skip++; continue; }
      if (idxByNorm[_rcN(nomMp)] || (skuMp && idxBySku[skuMp.toUpperCase()])) {
        stats.skip++;
        log.push('SKIP (ya existe): ' + nomMp);
        continue;
      }
      var newId  = skuMp || _rcMakeId(nomMp, secMp);
      var newRow = new Array(prodHdr.length).fill('');
      if (iID   >= 0) newRow[iID]   = newId;
      if (iSKU  >= 0) newRow[iSKU]  = skuMp;
      if (iDESC >= 0) newRow[iDESC] = nomMp;
      if (iCAT  >= 0) newRow[iCAT]  = secMp;
      if (iACT  >= 0) newRow[iACT]  = true;
      prodSh.appendRow(newRow);
      var newShRow = prodSh.getLastRow();
      idxByNorm[_rcN(nomMp)] = newShRow;
      if (skuMp) idxBySku[skuMp.toUpperCase()] = newShRow;
      stats.agregar++;
      log.push('AGREGADO: ' + nomMp + ' [' + newId + ']');
    }
  }

  SpreadsheetApp.flush();
  try { CacheService.getScriptCache().remove('erp_productos_v1'); } catch(e) {}

  Logger.log('\n=== CORRECCIONES APLICADAS ===');
  Logger.log('Actualizados (SKU/nombre/activo desde MP): ' + stats.actualizar);
  Logger.log('Reactivados (lista especial):              ' + stats.reactivar);
  Logger.log('Inactivados (confirmados):                 ' + stats.inactivar);
  Logger.log('Agregados nuevos desde MP:                 ' + stats.agregar);
  Logger.log('Omitidos:                                  ' + stats.skip);
  Logger.log('Errores (no encontrados):                  ' + stats.error);
  if (log.length) Logger.log('\nDetalle:\n' + log.join('\n'));
  Logger.log('\nBackup disponible en: ' + RC_BACKUP);
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function _rcN(s) {
  return String(s).toLowerCase()
    .replace(/[áàäâã]/g,'a').replace(/[éèëê]/g,'e')
    .replace(/[íìïî]/g,'i').replace(/[óòöôõ]/g,'o')
    .replace(/[úùüû]/g,'u').replace(/ñ/g,'n')
    .replace(/[^a-z0-9 ]/g,' ').replace(/\s+/g,' ').trim();
}

function _rcFuzzy(desc, mpList, minScore) {
  var norm = _rcN(desc);
  var ta   = norm.split(' ').filter(function(t){ return t.length > 2; });
  var best = null, bestScore = 0;
  for (var i = 0; i < mpList.length; i++) {
    var tb = _rcN(mpList[i].nombre).split(' ').filter(function(t){ return t.length > 2; });
    if (!ta.length || !tb.length) continue;
    var common = ta.filter(function(t){ return tb.indexOf(t) >= 0; }).length;
    var s = (2 * common) / (ta.length + tb.length);
    if (s > bestScore && s >= minScore) { bestScore = s; best = mpList[i]; }
  }
  return best;
}

function _rcMakeId(nombre, seccion) {
  var words  = _rcN(nombre).split(' ').filter(function(w){ return w.length > 2; });
  var prefix = (seccion || 'X').substring(0, 3).toUpperCase();
  var suffix = words.slice(0, 2).map(function(w){ return w.substring(0, 3); }).join('').toUpperCase();
  return prefix + '-' + suffix + '-' + Math.floor(Math.random()*900+100);
}
