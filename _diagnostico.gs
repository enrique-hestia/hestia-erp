// _diagnostico_catalogo.gs
// Ejecuta diagnosticarFuentes() y comparte el log completo

var _MP_ID  = '1qIc9x0JJrb0e2qVxrGiZqHA3CD36r0Wh';
var _CNT_ID = '1zH98lB4tJ5LA9FTCRqE1DaiFgGAjUFOg';
var _CAT_ID = '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY';

function diagnosticarFuentes() {
  // ── MERCADO PAGO ──────────────────────────────────────────
  var mpSS  = SpreadsheetApp.openById(_MP_ID);
  var mpSh  = mpSS.getSheets()[0];
  var mpAll = mpSh.getDataRange().getValues();

  Logger.log('=== MERCADO PAGO ===');
  Logger.log('Total filas: ' + mpAll.length);
  Logger.log('Total columnas: ' + (mpAll[0] ? mpAll[0].length : 0));
  Logger.log('--- Filas 1 a 10 (raw) ---');
  for (var i = 0; i < Math.min(10, mpAll.length); i++) {
    Logger.log('Fila ' + (i+1) + ': ' + JSON.stringify(mpAll[i].slice(0,12)));
  }

  // ── CNTADIGITAL ───────────────────────────────────────────
  var cntSS  = SpreadsheetApp.openById(_CNT_ID);
  var cntSh  = cntSS.getSheets()[0];
  var cntAll = cntSh.getDataRange().getValues();

  Logger.log('=== CNTADIGITAL ===');
  Logger.log('Total filas: ' + cntAll.length);
  Logger.log('Total columnas: ' + (cntAll[0] ? cntAll[0].length : 0));
  Logger.log('--- Fila 1 (encabezados) ---');
  Logger.log(JSON.stringify(cntAll[0]));
  Logger.log('--- Filas 2 a 5 (datos) ---');
  for (var i = 1; i < Math.min(5, cntAll.length); i++) {
    Logger.log('Fila ' + (i+1) + ': ' + JSON.stringify(cntAll[i]));
  }

  // ── CATÁLOGO GENERAL ──────────────────────────────────────
  var catSS  = SpreadsheetApp.openById(_CAT_ID);
  var catSh  = catSS.getSheets()[0];
  var catAll = catSh.getDataRange().getValues();

  Logger.log('=== CATÁLOGO GENERAL ===');
  Logger.log('Total filas: ' + catAll.length);
  Logger.log('--- Fila 1 (encabezados) ---');
  Logger.log(JSON.stringify(catAll[0]));
  Logger.log('--- Filas 2 a 6 (datos) ---');
  for (var i = 1; i < Math.min(6, catAll.length); i++) {
    Logger.log('Fila ' + (i+1) + ': ' + JSON.stringify(catAll[i]));
  }

  // ── COMPARACIÓN: primeros 5 nombres normalizados de cada fuente ──
  Logger.log('=== COMPARACIÓN DE NOMBRES NORMALIZADOS ===');
  Logger.log('--- MP (buscando fila donde empiezan los datos) ---');
  for (var i = 0; i < Math.min(mpAll.length, 10); i++) {
    var nombre = String(mpAll[i][1] || '').trim();
    var sku    = String(mpAll[i][4] || mpAll[i][0] || '').trim();
    if (nombre) Logger.log('MP fila ' + (i+1) + ': sku=' + sku + ' | nombre=' + nombre);
  }
  Logger.log('--- Catálogo (primeros 5 nombres) ---');
  for (var i = 1; i < Math.min(6, catAll.length); i++) {
    var desc = String(catAll[i][2] || '').trim();
    var sku  = String(catAll[i][1] || '').trim();
    if (desc) Logger.log('CAT fila ' + (i+1) + ': sku=' + sku + ' | desc=' + desc);
  }
}
