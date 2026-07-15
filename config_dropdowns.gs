var CFG_DD_TAB = 'Config_Dropdowns';

var CFG_DD_DEFAULTS = [
  ['Egresos',  'subtipo',    'Subtipo / Categoría', 'Honorarios Médicos|Honorarios Cons|Nómina|Renta|Servicios Generales|Medicamentos e Insumos|Insumos Lab|Laboratorio Externo|Marketing|Mantenimiento|Seguros|Impuestos y Contribuciones|Equipo Médico|Tecnología|Viáticos|Comisiones|Otros'],
  ['Egresos',  'contable',   'Tipo contable',        'Gasto|Costo|Crédito|Inversión'],
  ['Egresos',  'tipo',       'Tipo de egreso',       'Fijo|Variable|Extraordinario'],
  ['Egresos',  'formaPago',  'Forma de pago',        'Santander|Mercado Pago|AMEX|Efectivo|TDC|TDD|Transferencia'],
  ['CxP',      'subtipo',    'Subtipo / Categoría', 'Honorarios Médicos|Honorarios Cons|Nómina|Renta|Servicios Generales|Medicamentos e Insumos|Insumos Lab|Laboratorio Externo|Marketing|Mantenimiento|Seguros|Impuestos y Contribuciones|Equipo Médico|Tecnología|Viáticos|Comisiones|Otros'],
  ['CxP',      'formaPago',  'Forma de pago',        'Santander|Mercado Pago|AMEX|Efectivo|TDC|TDD|Transferencia'],
  ['Ingresos', 'formaPago',  'Forma de pago',        'Efectivo|Mercado Pago|AMEX|Santander|TDC|TDD|Transferencia|Cortesía'],
  ['Ingresos', 'facturacion','Facturación',           'No Factura|Factura|Factura Global|REPROVIDA|Grupo Médico'],
  ['General',  'sucursal',  'Sucursales',            'Lomas|Santa Fe|Interlomas'],
  ['General',  'moneda',    'Moneda',                'MXN|USD|EUR'],
];

function setupConfigDropdowns() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CFG_DD_TAB);
  if (!sh) {
    sh = ss.insertSheet(CFG_DD_TAB);
    sh.getRange(1,1,1,5).setValues([['Seccion','Campo','Etiqueta','Valores','Activo']]);
    sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#c46a7a').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  var data = sh.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) existing[data[i][0]+'|'+data[i][1]] = true;
  var toAdd = CFG_DD_DEFAULTS.filter(function(r){ return !existing[r[0]+'|'+r[1]]; })
                             .map(function(r){ return [r[0],r[1],r[2],r[3],true]; });
  if (toAdd.length) sh.getRange(sh.getLastRow()+1,1,toAdd.length,5).setValues(toAdd);
  sh.autoResizeColumns(1,5);
  return {ok:true, msg:'Config_Dropdowns lista. Filas añadidas: '+toAdd.length};
}

function readDropdowns() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(CFG_DD_TAB);
    if (!sh) { setupConfigDropdowns(); sh = ss.getSheetByName(CFG_DD_TAB); }
    var data = sh.getDataRange().getValues();
    var result = {};
    for (var i = 1; i < data.length; i++) {
      var seccion  = String(data[i][0]||'').trim();
      var campo    = String(data[i][1]||'').trim();
      var etiqueta = String(data[i][2]||'').trim();
      var valores  = String(data[i][3]||'').trim();
      var activo   = data[i][4]===true || String(data[i][4]).toUpperCase()==='TRUE';
      if (!seccion || !campo || !activo) continue;
      if (!result[seccion]) result[seccion] = {};
      result[seccion][campo] = {
        etiqueta: etiqueta,
        valores: valores ? valores.split('|').map(function(v){return v.trim();}).filter(Boolean) : []
      };
    }
    if (result['Egresos'] && result['Egresos']['subtipo']) {
      if (!result['CxP']) result['CxP'] = {};
      if (!result['CxP']['subtipo']) result['CxP']['subtipo'] = {etiqueta:'Subtipo / Categoría',valores:[]};
      result['CxP']['subtipo'].valores = result['Egresos']['subtipo'].valores.slice();
    }
    return {ok:true, dropdowns:result};
  } catch(ex) { return {ok:false, error:ex.message, dropdowns:{}}; }
}

function saveDropdownValues(body) {
  try {
    var seccion = String(body.seccion||'').trim();
    var campo   = String(body.campo  ||'').trim();
    var valores = String(body.valores||'').trim();
    if (!seccion || !campo) return {ok:false, error:'Seccion y Campo son requeridos'};
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(CFG_DD_TAB);
    if (!sh) { setupConfigDropdowns(); sh = ss.getSheetByName(CFG_DD_TAB); }
    var data = sh.getDataRange().getValues();
    var found = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim()===seccion && String(data[i][1]).trim()===campo) { found=i+1; break; }
    }
    if (found > 0) {
      sh.getRange(found, 4).setValue(valores);
    } else {
      var etiqueta = body.etiqueta || (campo.charAt(0).toUpperCase()+campo.slice(1));
      sh.appendRow([seccion, campo, etiqueta, valores, true]);
    }
    if (seccion==='Egresos' && campo==='subtipo') {
      var ds = sh.getDataRange().getValues(); var cxpRow=-1;
      for (var j=1;j<ds.length;j++) { if(String(ds[j][0]).trim()==='CxP'&&String(ds[j][1]).trim()==='subtipo'){cxpRow=j+1;break;} }
      if (cxpRow>0) sh.getRange(cxpRow,4).setValue(valores);
      else sh.appendRow(['CxP','subtipo','Subtipo / Categoría',valores,true]);
    }
    SpreadsheetApp.flush();
    return {ok:true, seccion:seccion, campo:campo};
  } catch(ex) { return {ok:false, error:ex.message}; }
}