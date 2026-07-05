// ============================================================
// FACTURACIÓN — Conciliación XML (Drive) vs Ingresos + Reporte ContaDigital
// ============================================================
// Estructura de carpetas en Drive (fija, ver memoria del proyecto):
// .../onefactureXMLs/HCL2307051Y6/emitidos/{año}/{MM MES}/  (archivos UUID.xml)

var FAC_MESES_ABR = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];

// Encabezados EXACTOS (texto, acentos y espacios) copiados de la plantilla real
// ContaDigital "Plantilla_Masiva.xlsx" que usa el usuario — no cambiar sin volver
// a comparar contra esa plantilla, porque el importador de ContaDigital hace
// match por nombre de columna literal.
var FAC_MASIVA_HEADERS_INV = ['SERIE','FECHA','SUCURSAL','RFC RECEPTOR','MONEDA','T.C.','MÉTODO DE PAGO','CUENTA BANCARIA','REFERENCIA','IMPORTE','DESCUENTO','SUBTOTAL','IVA TOTAL','IEPS TOTAL','RETENCIÓN IVA','RETENCIÓN ISR','TOTAL FACTURA '];
var FAC_MASIVA_HEADERS_CONCEPTO = ['CLAVE PRODUCTO','DESCRIPCION ADICIONAL ','CANTIDAD ','PRECIO ','DESCUENTO $','IVA % ',' IVA $ ','IEPS %','IEPS $ ','RET IVA% ','RET IVA $','RET ISR % ','RET ISR $','PREDIAL'];
var FAC_MASIVA_HEADERS_TRAIL = ['USO CFDI','CONDICIONES DE PAGO','FORMA DE PAGO','TIPO','ALUMNO','CURP','NIVEL','RVOE','TIPO RELACION','UUID RELACIONADOS','FECHA VENCIMIENTO'];
var FAC_MASIVA_CONCEPTO_BLOQUES = 10; // la plantilla real siempre trae 10 bloques de concepto, aunque la factura tenga menos líneas
var FAC_MASIVA_SERIE = 'Hestia Pacientes'; // constante fija de la plantilla, no viene del CFDI

// Carpeta de Drive donde el usuario pidió que se guarde siempre el reporte masivo
// (no la carpeta genérica de facturas de Ingresos).
var FAC_MASIVA_FOLDER_ID = '1yf0a6NTIMJdDpTdqO6SRgTPKVcuIoONJ';

var FAC_METODO_PAGO_MAP = {
  '01':'Efectivo','02':'Cheque nominativo','03':'Transferencia electrónica de fondos',
  '04':'Tarjeta de crédito','05':'Monedero electrónico','06':'Dinero electrónico',
  '08':'Vales de despensa','28':'Tarjeta de débito','29':'Tarjeta de servicios',
  '31':'Tarjeta de crédito','99':'Tarjeta de crédito'
};

var FAC_MONEDA_MAP = { 'MXN':'PESOS', 'USD':'USD', 'EUR':'EUR', 'CAD':'CAD' };

// Construye un .xlsx válido (OOXML) directamente con Utilities.zip(), sin pasar
// por SpreadsheetApp/DriveApp.getAs()/UrlFetchApp — evita las conversiones de
// Drive que fallan de forma intermitente y no requiere ningún permiso adicional.
function _buildXlsxBlob(rows, sheetName, fileName) {
  function colLetter(n) {
    var s = ''; n++;
    while (n > 0) { var rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function esc(v) {
    return String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  }

  var sheetRowsXml = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r] || [];
    var cellsXml = [];
    for (var c = 0; c < row.length; c++) {
      var v = row[c];
      if (v === '' || v === null || v === undefined) continue;
      var ref = colLetter(c) + (r + 1);
      if (typeof v === 'number' && isFinite(v)) {
        cellsXml.push('<c r="' + ref + '"><v>' + v + '</v></c>');
      } else {
        cellsXml.push('<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>');
      }
    }
    sheetRowsXml.push('<row r="' + (r + 1) + '">' + cellsXml.join('') + '</row>');
  }

  var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    + '</Types>';

  var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    + '</Relationships>';

  var workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<sheets><sheet name="' + esc(sheetName) + '" sheetId="1" r:id="rId1"/></sheets>'
    + '</workbook>';

  var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + '</Relationships>';

  var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    + '</styleSheet>';

  var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<sheetData>' + sheetRowsXml.join('') + '</sheetData>'
    + '</worksheet>';

  var parts = [
    Utilities.newBlob(contentTypes, 'application/xml', '[Content_Types].xml'),
    Utilities.newBlob(rootRels, 'application/xml', '_rels/.rels'),
    Utilities.newBlob(workbookXml, 'application/xml', 'xl/workbook.xml'),
    Utilities.newBlob(workbookRels, 'application/xml', 'xl/_rels/workbook.xml.rels'),
    Utilities.newBlob(stylesXml, 'application/xml', 'xl/styles.xml'),
    Utilities.newBlob(sheetXml, 'application/xml', 'xl/worksheets/sheet1.xml')
  ];

  var zipBlob = Utilities.zip(parts, fileName);
  return Utilities.newBlob(zipBlob.getBytes(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
}

function _facMonthFolder(anio, mes) {
  var mesTag = (mes < 10 ? '0' : '') + mes + ' ' + FAC_MESES_ABR[mes - 1];
  var top = DriveApp.getFoldersByName('onefactureXMLs');
  if (!top.hasNext()) return null;
  var f = top.next();
  var chain = ['HCL2307051Y6', 'emitidos', String(anio), mesTag];
  for (var i = 0; i < chain.length; i++) {
    var sub = f.getFoldersByName(chain[i]);
    if (!sub.hasNext()) return null;
    f = sub.next();
  }
  return f;
}

function _facMonthsInRange(fechaInicio, fechaFin) {
  var out = [];
  var d0 = new Date(fechaInicio + 'T00:00:00');
  var d1 = new Date(fechaFin + 'T00:00:00');
  var y = d0.getFullYear(), m = d0.getMonth() + 1;
  var yEnd = d1.getFullYear(), mEnd = d1.getMonth() + 1;
  var guard = 0;
  while ((y < yEnd || (y === yEnd && m <= mEnd)) && guard < 36) {
    out.push({ anio: y, mes: m });
    m++; if (m > 12) { m = 1; y++; }
    guard++;
  }
  return out;
}

// Igual que _facMonthsInRange pero agrega N meses colindantes por delante y por
// detrás — para encontrar CFDI de ingresos del periodo que se timbraron en el
// mes anterior/siguiente (el caso real: venta de junio facturada el 2 de julio).
function _facMonthsInRangePadded(fechaInicio, fechaFin, pad) {
  pad = pad || 1;
  var base = _facMonthsInRange(fechaInicio, fechaFin);
  if (!base.length) return base;
  var out = base.slice();
  var first = base[0], last = base[base.length - 1];
  for (var i = 1; i <= pad; i++) {
    var mA = first.mes - i, yA = first.anio;
    while (mA < 1) { mA += 12; yA--; }
    out.unshift({ anio: yA, mes: mA, colindante: true });
    var mB = last.mes + i, yB = last.anio;
    while (mB > 12) { mB -= 12; yB++; }
    out.push({ anio: yB, mes: mB, colindante: true });
  }
  return out;
}

// Rango de fechas de los meses colindantes (antes y después) al periodo
// analizado — para poder buscar operaciones pendientes (sin folio) de un mes
// cercano antes de declarar un XML como "sobrante". Devuelve fechas ISO del
// primer y último día de ese rango ampliado.
function _facRangoVecino(fechaInicio, fechaFin, pad) {
  try {
    pad = pad || 1;
    var d0 = new Date(fechaInicio + 'T00:00:00');
    var d1 = new Date(fechaFin + 'T00:00:00');
    var start = new Date(d0.getFullYear(), d0.getMonth() - pad, 1);
    var end = new Date(d1.getFullYear(), d1.getMonth() + pad + 1, 0); // día 0 del mes siguiente = último día del mes objetivo
    function iso(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
    return { inicio: iso(start), fin: iso(end) };
  } catch (e) { return null; }
}

// Normaliza un folio para matching: "HF 1164", "hf-1164", " 1164 " y "01164"
// deben encontrar el mismo XML. Se queda con la serie alfabética (si la hay)
// + el número sin ceros a la izquierda.
function _facNormFolio(s) {
  var t = String(s == null ? '' : s).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  var m = t.match(/^([A-Z]*)0*(\d+)$/);
  if (m) return m[2]; // solo el número — la serie del emisor es única (HF), no discrimina
  return t;
}

function _facQuickParse(xml) {
  var headEnd = xml.indexOf('<cfdi:Emisor');
  var head = headEnd > -1 ? xml.substring(0, headEnd) : xml;
  function a(name, src) { var m = (src || head).match(new RegExp(name + '="([^"]*)"')); return m ? m[1] : ''; }
  var recMatch = xml.match(/<cfdi:Receptor\s+([^>]*)\/>/);
  var rec = recMatch ? recMatch[1] : '';
  var uuidM = xml.match(/UUID="([^"]*)"/);
  return {
    folio: a('Folio'), serie: a('Serie'), tipo: a('TipoDeComprobante'),
    total: parseFloat(a('Total')) || 0, subTotal: parseFloat(a('SubTotal')) || 0,
    descuento: parseFloat(a('Descuento')) || 0,
    fecha: a('Fecha'), formaPago: a('FormaPago'),
    receptorNombre: a('Nombre', rec), receptorRfc: a('Rfc', rec),
    receptorCP: a('DomicilioFiscalReceptor', rec), receptorUsoCfdi: a('UsoCFDI', rec),
    receptorRegimen: a('RegimenFiscalReceptor', rec),
    uuid: uuidM ? uuidM[1] : ''
  };
}

function _facBuildXmlIndex(fechaInicio, fechaFin) {
  // Se escanean también los meses colindantes (±1) y NO se descarta ningún
  // CFDI tipo I por fecha: un ingreso del periodo pudo haberse timbrado en
  // otro mes, y eso es exactamente lo que hay que detectar (no esconder).
  // Cada XML lleva enPeriodo (su fecha cae dentro del rango analizado) y
  // mes (YYYY-MM del timbrado) para poder clasificar cruces de periodo.
  var months = _facMonthsInRangePadded(fechaInicio, fechaFin, 1);
  var byFolio = {}, byUuid = {}, byFileId = {}, all = [], carpetas = [];
  months.forEach(function (ym) {
    var mesTag = (ym.mes < 10 ? '0' : '') + ym.mes + ' ' + FAC_MESES_ABR[ym.mes - 1];
    var pathStr = 'onefactureXMLs/HCL2307051Y6/emitidos/' + ym.anio + '/' + mesTag;
    var folder = _facMonthFolder(ym.anio, ym.mes);
    var totalArchivos = 0, totalIngreso = 0;
    if (folder) {
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var name = file.getName();
        if (!/\.xml$/i.test(name)) continue;
        totalArchivos++;
        var xml;
        try { xml = file.getBlob().getDataAsString('UTF-8'); } catch (e) { continue; }
        var p = _facQuickParse(xml);
        if (p.tipo !== 'I') continue; // solo facturas de ingreso (excluye nómina, notas de crédito, pago)
        var fechaISO = p.fecha ? p.fecha.substring(0, 10) : '';
        totalIngreso++;
        var rec = {
          folio: p.folio, serie: p.serie, uuid: p.uuid, total: p.total,
          subTotal: p.subTotal, descuento: p.descuento, fecha: fechaISO,
          mes: fechaISO.substring(0, 7),
          enPeriodo: (fechaISO >= fechaInicio && fechaISO <= fechaFin),
          receptorNombre: p.receptorNombre, receptorRfc: p.receptorRfc,
          receptorCP: p.receptorCP, receptorUsoCfdi: p.receptorUsoCfdi, receptorRegimen: p.receptorRegimen,
          formaPago: p.formaPago,
          fileId: file.getId(), fileUrl: file.getUrl(), fileName: name
        };
        all.push(rec);
        if (p.folio) {
          byFolio[String(p.folio)] = rec;
          var nf = _facNormFolio(p.folio);
          if (nf && !byFolio[nf]) byFolio[nf] = rec;
        }
        if (p.uuid) byUuid[p.uuid.toUpperCase()] = rec;
        byFileId[rec.fileId] = rec;
      }
    }
    carpetas.push({ path: pathStr, encontrada: !!folder, totalArchivosXml: totalArchivos, totalTipoIngreso: totalIngreso, colindante: !!ym.colindante });
  });
  return { byFolio: byFolio, byUuid: byUuid, byFileId: byFileId, all: all, carpetas: carpetas };
}

// Busca en el índice el XML de una operación por folio normalizado o UUID.
function _facXmlPorOp(idx, op) {
  if (op.facturaUUID && idx.byUuid[String(op.facturaUUID).toUpperCase()]) return idx.byUuid[String(op.facturaUUID).toUpperCase()];
  if (op.factura) {
    if (idx.byFolio[op.factura]) return idx.byFolio[op.factura];
    var nf = _facNormFolio(op.factura);
    if (nf && idx.byFolio[nf]) return idx.byFolio[nf];
  }
  return null;
}

// Detecta la porción pagada con Nota de Crédito: es un descuento (saldo a
// favor del paciente), NO un cobro — no se factura ni entra a la declaración.
function _facEsNotaCredito(fp) {
  return String(fp || '').trim().toLowerCase().indexOf('nota de cr') === 0;
}

function _facReadOpsInRange(fechaInicio, fechaFin) {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sheet = ss.getSheetByName(BD_INGRESOS_TAB);
  if (!sheet) return [];
  var raw = sheet.getDataRange().getValues();
  // Columnas dinámicas (agregadas por migración, viven al final de la hoja)
  var hdrs = (raw[0] || []).map(function (h) { return String(h).trim().toLowerCase(); });
  var iUuid = hdrs.indexOf('facturauuid'), iRfc = hdrs.indexOf('facturarfc'), iRazon = hdrs.indexOf('razonsocial');
  var iPagosDet = hdrs.indexOf('pagosdetalle');
  var opsMap = {}, order = [];
  function dt(v) {
    if (!v) return '';
    if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    return String(v).substring(0, 10);
  }
  function num(v) { var n = parseFloat(String(v || '').replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var op = String(r[0] || '').trim();
    if (!op) continue;
    var fecha = dt(r[2]);
    if (fecha < fechaInicio || fecha > fechaFin) continue;
    if (!opsMap[op]) {
      opsMap[op] = {
        id: op, fecha: fecha, paciente: String(r[3] || ''), total: 0,
        factura: '', poliza: '', archivoURL: '', facturaUUID: '', facturaRFC: '', razonSocial: '',
        formaPago: '', notaCredito: 0
      };
      order.push(op);
    }
    var o = opsMap[op];
    o.total += num(r[9]);
    // Factura#/Póliza/ArchivoURL/UUID viven en UNA sola línea de la OP (no
    // necesariamente la primera) — se fusionan de todas las líneas. Leerlos
    // solo de la primera línea fue el bug que marcaba como "pendiente por
    // facturar" operaciones ya facturadas (y provocó facturación doble).
    if (!o.factura && String(r[17] || '').trim()) o.factura = String(r[17]).trim();
    if (!o.poliza && String(r[18] || '').trim()) o.poliza = String(r[18]).trim();
    if (!o.archivoURL && String(r[22] || '').trim()) o.archivoURL = String(r[22]).trim();
    if (!o.formaPago && String(r[12] || '').trim()) o.formaPago = String(r[12]).trim();
    if (iUuid > -1 && !o.facturaUUID && String(r[iUuid] || '').trim()) o.facturaUUID = String(r[iUuid]).trim();
    if (iRfc > -1 && !o.facturaRFC && String(r[iRfc] || '').trim()) o.facturaRFC = String(r[iRfc]).trim();
    if (iRazon > -1 && !o.razonSocial && String(r[iRazon] || '').trim()) o.razonSocial = String(r[iRazon]).trim();
    // Pago mixto: la porción de Nota de Crédito viene del desglose PagosDetalle
    if (iPagosDet > -1 && !o._pagosDetLeido && String(r[iPagosDet] || '').trim()) {
      o._pagosDetLeido = true;
      try {
        JSON.parse(String(r[iPagosDet])).forEach(function (p) {
          if (_facEsNotaCredito(p.fp)) o.notaCredito += parseFloat(p.monto) || 0;
        });
      } catch (e) {}
    }
  }
  // Op pagada COMPLETA con Nota de Crédito (sin desglose): todo es descuento
  order.forEach(function (k) {
    var o = opsMap[k];
    if (!o.notaCredito && _facEsNotaCredito(o.formaPago)) o.notaCredito = o.total;
    o.totalFacturable = Math.max(0, o.total - o.notaCredito);
    delete o._pagosDetLeido;
  });
  return order.map(function (k) { return opsMap[k]; });
}

// Igual que _facReadOpsInRange pero además suma el precio de lista (PVP × Cantidad,
// SIN descuento) y detecta si alguna línea de la venta llevó % de descuento —
// necesario para el análisis de descuento fiscal vs. descuento en la venta.
function _facReadOpsConDetalleVenta(fechaInicio, fechaFin) {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sheet = ss.getSheetByName(BD_INGRESOS_TAB);
  if (!sheet) return [];
  var raw = sheet.getDataRange().getValues();
  var opsMap = {}, order = [];
  function dt(v) {
    if (!v) return '';
    if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    return String(v).substring(0, 10);
  }
  function num(v) { var n = parseFloat(String(v || '').replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var op = String(r[0] || '').trim();
    if (!op) continue;
    var fecha = dt(r[2]);
    if (fecha < fechaInicio || fecha > fechaFin) continue;
    if (!opsMap[op]) {
      opsMap[op] = {
        id: op, fecha: fecha, paciente: String(r[3] || ''), total: 0, montoLista: 0,
        tieneDescuentoVenta: false, descuentoVentaPct: 0,
        factura: '', archivoURL: ''
      };
      order.push(op);
    }
    // Igual que _facReadOpsInRange: factura/URL pueden vivir en cualquier línea
    if (!opsMap[op].factura && String(r[17] || '').trim()) opsMap[op].factura = String(r[17]).trim();
    if (!opsMap[op].archivoURL && String(r[22] || '').trim()) opsMap[op].archivoURL = String(r[22]).trim();
    var pvp = num(r[6]), descPct = num(r[7]), cant = num(r[8]) || 1;
    opsMap[op].total += num(r[9]);
    opsMap[op].montoLista += pvp * cant;
    if (descPct > 0) {
      opsMap[op].tieneDescuentoVenta = true;
      if (descPct > opsMap[op].descuentoVentaPct) opsMap[op].descuentoVentaPct = descPct;
    }
  }
  return order.map(function (k) { return opsMap[k]; });
}

// Lee BD_Ingresos con detalle por línea (Producto/PVP/Descuento%/Cantidad) SOLO
// para los OP indicados — usado por el reporte de "Pendiente por facturar", que
// arma los bloques de concepto de ContaDigital a partir de la venta real (no hay
// XML del que leer porque, por definición, todavía no se ha facturado).
function _facReadOpsPendientesDetalle(fechaInicio, fechaFin, opIds) {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sheet = ss.getSheetByName(BD_INGRESOS_TAB);
  if (!sheet) return {};
  var raw = sheet.getDataRange().getValues();
  var wanted = {};
  opIds.forEach(function (id) { wanted[id] = true; });
  function dt(v) {
    if (!v) return '';
    if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    return String(v).substring(0, 10);
  }
  function num(v) { var n = parseFloat(String(v || '').replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
  var opsMap = {};
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var op = String(r[0] || '').trim();
    if (!op || !wanted[op]) continue;
    var fecha = dt(r[2]);
    if (fecha < fechaInicio || fecha > fechaFin) continue;
    if (!opsMap[op]) {
      opsMap[op] = {
        id: op, fecha: fecha, paciente: String(r[3] || ''), sucursal: String(r[21] || ''),
        formaPago: String(r[12] || ''), montoLista: 0, total: 0, lineas: []
      };
    }
    var pvp = num(r[6]), descPct = num(r[7]), cant = num(r[8]) || 1, totalLinea = num(r[9]);
    opsMap[op].lineas.push({ producto: String(r[5] || ''), cantidad: cant, pvp: pvp, descPct: descPct });
    opsMap[op].montoLista += pvp * cant;
    opsMap[op].total += totalLinea;
  }
  return opsMap;
}

function reconciliarFacturasXml(fechaInicio, fechaFin) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    var ops = _facReadOpsInRange(fechaInicio, fechaFin);
    var idx = _facBuildXmlIndex(fechaInicio, fechaFin);
    var conDocumento = [], faltaDocumento = [], sinFactura = [], notasCredito = [];
    var usedFolios = {}, usedFileIds = {};
    // Operaciones SIN folio capturado de los meses colindantes (no del periodo
    // exacto — esas ya están en "ops") — antes de declarar un XML como
    // "sobrante" (candidato a cancelar en el SAT) hay que descartar que en
    // realidad sea la factura de una venta de un mes cercano que todavía no
    // capturó su folio. Sin este cruce, un CFDI que sí tiene venta detrás
    // podía marcarse como exceso por pura coincidencia de fechas de análisis.
    var opsVecinasSinFactura = [];
    var rangoVecino = _facRangoVecino(fechaInicio, fechaFin, 1);
    if (rangoVecino) {
      var opsEnPeriodo = {};
      ops.forEach(function (o) { opsEnPeriodo[o.id] = true; });
      _facReadOpsInRange(rangoVecino.inicio, rangoVecino.fin).forEach(function (o) {
        if (!opsEnPeriodo[o.id] && !o.factura) opsVecinasSinFactura.push(o);
      });
    }
    // Dueños de facturas en TODO BD_Ingresos (cualquier fecha): un XML que ya
    // está asignado a alguna operación — del periodo o de otro mes — NUNCA
    // debe ofrecerse como sugerencia/candidato para otra operación. Que el
    // folio 1081 (ya asignado a una OP de junio) apareciera como candidato
    // para las OPs de julio era exactamente el tipo de confusión que causa
    // dobles vinculaciones.
    var duenoIdx = _facIndexDuenosBDIngresos();
    function estaAsignado(x) {
      return !!(duenoIdx.porFileId[x.fileId]
        || (x.uuid && duenoIdx.porUuid[x.uuid.toUpperCase()])
        || (x.folio && duenoIdx.porFolio[_facNormFolio(x.folio)]));
    }
    function marcarUsado(x) {
      if (!x) return;
      usedFileIds[x.fileId] = true;
      if (x.folio) { usedFolios[x.folio] = true; usedFolios[_facNormFolio(x.folio)] = true; }
    }
    // Cruce de periodo: el CFDI se declara en el MES en que se timbró — si ese
    // mes no es el mes de la venta, hay una diferencia fiscal que reportar.
    function tagOtroMes(target, op, x) {
      if (!x || !x.fecha) return;
      var opMes = String(op.fecha || '').substring(0, 7);
      if (x.mes && opMes && x.mes !== opMes) {
        target.xmlOtroMes = true;
        target.xmlFecha = x.fecha;
        target.xmlMes = x.mes;
      }
    }
    ops.forEach(function (op) {
      if (op.archivoURL) {
        // Localizar el XML real detrás de la URL para: (a) que no aparezca
        // como huérfano, (b) detectar si se timbró en otro mes.
        var mUrl = op.archivoURL.match(/[-\w]{25,}/);
        var xLinked = mUrl ? idx.byFileId[mUrl[0]] : null;
        if (!xLinked) xLinked = _facXmlPorOp(idx, op);
        marcarUsado(xLinked);
        if (xLinked) { tagOtroMes(op, op, xLinked); op.xmlFolio = xLinked.folio; }
        conDocumento.push(op);
        return;
      }
      var x = _facXmlPorOp(idx, op);
      if (x) {
        marcarUsado(x);
        var fd = {
          id: op.id, fecha: op.fecha, paciente: op.paciente, total: op.total, factura: op.factura,
          notaCredito: op.notaCredito || 0, totalFacturable: op.totalFacturable,
          xmlFileId: x.fileId, xmlFileUrl: x.fileUrl, xmlTotal: x.total, xmlUuid: x.uuid,
          xmlRazonSocial: x.receptorNombre, xmlRfc: x.receptorRfc
        };
        tagOtroMes(fd, op, x);
        faltaDocumento.push(fd);
        return;
      }
      // Pagada completa con Nota de Crédito y sin CFDI: no es "pendiente por
      // facturar" — es un descuento (no fiscal). Se lista aparte para que la
      // conciliación total cuadre sin confundirla con un faltante real.
      if ((op.totalFacturable || 0) <= 0.01 && (op.notaCredito || 0) > 0) {
        notasCredito.push(op);
        return;
      }
      sinFactura.push(op);
    });

    // Para las que no tienen folio capturado: sugerir el XML más probable por
    // nombre de paciente (+ monto como desempate), sin necesidad de que el
    // usuario ya sepa/escriba el número de factura primero.
    sinFactura.forEach(function (op) {
      var norm = _pacNormNombre(op.paciente);
      if (!norm) return;
      // Muchos pacientes se capturan como pareja ("Aleksandra y Kevin") pero el
      // XML solo trae el nombre de uno — se prueba también cada nombre por separado.
      var candidatos = [norm];
      norm.split(' y ').forEach(function (part) {
        part = part.trim();
        if (part.length >= 3 && candidatos.indexOf(part) === -1) candidatos.push(part);
      });
      var mejor = null, mejorDelta = Infinity;
      for (var i = 0; i < idx.all.length; i++) {
        var x = idx.all[i];
        if (usedFileIds[x.fileId] || estaAsignado(x)) continue;
        var xNorm = _pacNormNombre(x.receptorNombre);
        if (!xNorm) continue;
        var matched = candidatos.some(function (c) { return c && (xNorm.indexOf(c) > -1 || c.indexOf(xNorm) > -1); });
        if (!matched) continue;
        // El CFDI se timbra por lo realmente cobrado — la porción de Nota de
        // Crédito es descuento, así que el monto a comparar es el facturable.
        var montoOp = (op.totalFacturable !== undefined) ? op.totalFacturable : (op.total || 0);
        var delta = Math.abs((x.total || 0) - montoOp);
        if (delta < mejorDelta) { mejor = x; mejorDelta = delta; }
      }
      if (mejor) {
        op.sugerenciaFolio = mejor.folio;
        op.sugerenciaFileId = mejor.fileId;
        op.sugerenciaTotal = mejor.total;
        op.sugerenciaFecha = mejor.fecha;
        op.sugerenciaCoincideMonto = mejorDelta < 0.01;
        op.sugerenciaRazonSocial = mejor.receptorNombre;
        op.sugerenciaRfc = mejor.receptorRfc;
        op.sugerenciaUuid = mejor.uuid;
        op.sugerenciaTipo = 'nombre';
        tagOtroMes(op, op, mejor);
        marcarUsado(mejor); // no ofrecer el mismo XML como sugerencia a dos operaciones
      }
    });

    // Segunda pasada, solo para lo que sigue sin sugerencia: muchas facturas se
    // timbran a "PÚBLICO EN GENERAL" (sin nombre del paciente en el XML), así
    // que no hay con qué comparar por nombre — se ofrece por monto exacto
    // dentro del mismo periodo analizado. Si hay más de un XML con el mismo
    // monto no se autoselecciona ninguno (riesgo de vincular el documento
    // fiscal equivocado) — se listan todos como candidatos para que el
    // usuario elija a mano.
    sinFactura.forEach(function (op) {
      if (op.sugerenciaFolio) return;
      // Un XML timbrado a nombre de OTRA persona no es candidato automático por
      // monto (pertenece a quien dice su razón social; si de verdad es de esta
      // operación —ej. pagó el esposo— se vincula a mano). Solo los genéricos
      // (PÚBLICO EN GENERAL / RFC XAXX010101000) se sugieren por monto.
      function esGenerico(x) {
        var rfc = String(x.receptorRfc || '').toUpperCase();
        var nom = _pacNormNombre(x.receptorNombre);
        return rfc === 'XAXX010101000' || nom.indexOf('publico en general') > -1 || !nom;
      }
      var montoOpFac = (op.totalFacturable !== undefined) ? op.totalFacturable : (op.total || 0);
      var todosMonto = idx.all.filter(function (x) {
        return x.folio && !usedFileIds[x.fileId] && !estaAsignado(x) && Math.abs((x.total || 0) - montoOpFac) < 0.01;
      });
      var candidatos = todosMonto.filter(esGenerico);
      if (!candidatos.length) {
        // Sin genérico disponible: los con nombre se listan solo como candidatos
        // manuales, nunca autoseleccionados.
        if (todosMonto.length) {
          op.candidatosPorMonto = todosMonto.slice(0, 8).map(function (x) {
            return { folio: x.folio, fileId: x.fileId, fecha: x.fecha, total: x.total, razonSocial: x.receptorNombre, rfc: x.receptorRfc };
          });
        }
        return;
      }
      if (candidatos.length === 1) {
        var u = candidatos[0];
        op.sugerenciaFolio = u.folio;
        op.sugerenciaFileId = u.fileId;
        op.sugerenciaTotal = u.total;
        op.sugerenciaFecha = u.fecha;
        op.sugerenciaCoincideMonto = true;
        op.sugerenciaRazonSocial = u.receptorNombre;
        op.sugerenciaRfc = u.receptorRfc;
        op.sugerenciaUuid = u.uuid;
        op.sugerenciaTipo = 'monto';
        tagOtroMes(op, op, u);
        marcarUsado(u);
      } else {
        op.candidatosPorMonto = candidatos.slice(0, 8).map(function (x) {
          return { folio: x.folio, fileId: x.fileId, fecha: x.fecha, total: x.total, razonSocial: x.receptorNombre, rfc: x.receptorRfc };
        });
      }
    });

    // ── XML sin operación en el periodo ────────────────────────────────
    // Solo los timbrados DENTRO del rango cuentan aquí (los de meses
    // colindantes solo sirvieron para encontrar CFDIs de ingresos del
    // periodo). Antes de declarar un XML "sobrante", se busca su dueño en
    // TODO BD_Ingresos sin filtro de fecha: si está vinculado a una
    // operación de otro mes, no es exceso — es un CFDI cruzado de periodo.
    var xmlHuerfanos = idx.all.filter(function (x) { return x.enPeriodo && !usedFileIds[x.fileId]; });

    function esGenericoXml(x) {
      var rfc = String(x.receptorRfc || '').toUpperCase();
      var nom = _pacNormNombre(x.receptorNombre);
      return rfc === 'XAXX010101000' || nom.indexOf('publico en general') > -1 || !nom;
    }
    var vecinosUsados = {};
    xmlHuerfanos.forEach(function (x) {
      var dueno = duenoIdx.porFileId[x.fileId]
        || (x.uuid && duenoIdx.porUuid[x.uuid.toUpperCase()])
        || (x.folio && duenoIdx.porFolio[_facNormFolio(x.folio)]);
      if (dueno) {
        x.clasificacion = 'otroPeriodo';
        x.opDuena = dueno.opId;
        x.opDuenaFecha = dueno.fecha;
        x.opDuenaPaciente = dueno.paciente;
        return;
      }
      // Antes de declarar "sobrante" (candidato a cancelar en el SAT): ¿su
      // monto coincide con una venta pendiente (sin folio) de un mes cercano
      // que simplemente no ha capturado su factura todavía? Solo se cruza
      // contra XML genéricos (público en general) — uno con nombre real
      // pertenece a quien dice esa razón social, no se reasigna por monto.
      if (esGenericoXml(x)) {
        var vecino = opsVecinasSinFactura.find(function (o) {
          if (vecinosUsados[o.id]) return false;
          var montoOp = (o.totalFacturable !== undefined) ? o.totalFacturable : (o.total || 0);
          return Math.abs(montoOp - (x.total || 0)) < 0.01;
        });
        if (vecino) {
          x.clasificacion = 'pendienteVecino';
          x.vecinoOpId = vecino.id;
          x.vecinoFecha = vecino.fecha;
          x.vecinoPaciente = vecino.paciente;
          vecinosUsados[vecino.id] = true;
          return;
        }
      }
      x.clasificacion = 'sobrante';
    });

    // Posible duplicado: el monto coincide con una operación que YA tiene su
    // propia factura — señal fuerte de que se facturó dos veces la misma venta.
    var opsFacturadas = conDocumento.concat(faltaDocumento);
    xmlHuerfanos.forEach(function (x) {
      if (x.clasificacion === 'otroPeriodo') return;
      var match = opsFacturadas.find(function (op) { return Math.abs((op.total || 0) - (x.total || 0)) < 0.01; });
      if (match) { x.posibleDuplicadoDeOp = match.id; x.posibleDuplicadoDePaciente = match.paciente; }
    });

    // ── Totales ─────────────────────────────────────────────────────────
    var enPeriodo = idx.all.filter(function (x) { return x.enPeriodo; });
    var totalFacturadoXml = enPeriodo.reduce(function (s, x) { return s + (x.total || 0); }, 0);
    var totalIngresos = ops.reduce(function (s, op) { return s + (op.total || 0); }, 0);
    var sobrantes = xmlHuerfanos.filter(function (x) { return x.clasificacion === 'sobrante'; });
    var huerfanosOtroPeriodo = xmlHuerfanos.filter(function (x) { return x.clasificacion === 'otroPeriodo'; });
    var pendientesVecinos = xmlHuerfanos.filter(function (x) { return x.clasificacion === 'pendienteVecino'; });
    var totalHuerfanos = xmlHuerfanos.reduce(function (s, x) { return s + (x.total || 0); }, 0);
    var totalSobrantes = sobrantes.reduce(function (s, x) { return s + (x.total || 0); }, 0);
    var totalXmlDeOtroPeriodo = huerfanosOtroPeriodo.reduce(function (s, x) { return s + (x.total || 0); }, 0);
    var totalPendientesVecinos = pendientesVecinos.reduce(function (s, x) { return s + (x.total || 0); }, 0);

    // Ingresos del periodo cuyo CFDI se timbró en otro mes (confirmados o con
    // folio localizado — no sugerencias sin confirmar): van a la declaración
    // del mes del timbrado, no de este. faltaDocumento significa que el folio
    // ya está capturado pero el ArchivoURL todavía no — se marca para que la
    // sección de cruzados pueda ofrecer el botón "Vincular" ahí mismo.
    var opsFacturadasOtroMes = [];
    conDocumento.forEach(function (op) { if (op.xmlOtroMes) opsFacturadasOtroMes.push({ opId: op.id, paciente: op.paciente, fecha: op.fecha, total: op.total, folio: op.xmlFolio || op.factura, xmlFecha: op.xmlFecha, xmlMes: op.xmlMes, necesitaVincular: false }); });
    faltaDocumento.forEach(function (fd) { if (fd.xmlOtroMes) opsFacturadasOtroMes.push({ opId: fd.id, paciente: fd.paciente, fecha: fd.fecha, total: fd.total, folio: fd.factura, xmlFecha: fd.xmlFecha, xmlMes: fd.xmlMes, necesitaVincular: true, xmlFileId: fd.xmlFileId }); });
    var totalOpsFacturadasOtroMes = opsFacturadasOtroMes.reduce(function (s, o) { return s + (o.total || 0); }, 0);

    var totalNotasCredito = notasCredito.reduce(function (s, o) { return s + (o.notaCredito || 0); }, 0)
      + sinFactura.reduce(function (s, o) { return s + (o.notaCredito || 0); }, 0)
      + conDocumento.reduce(function (s, o) { return s + (o.notaCredito || 0); }, 0)
      + faltaDocumento.reduce(function (s, o) { return s + (o.notaCredito || 0); }, 0);

    return {
      ok: true, conDocumento: conDocumento, faltaDocumento: faltaDocumento,
      sinFactura: sinFactura, notasCredito: notasCredito, xmlHuerfanos: xmlHuerfanos, totalOps: ops.length,
      carpetasAnalizadas: idx.carpetas,
      cfdiCruzados: {
        opsFacturadasOtroMes: opsFacturadasOtroMes,        // ingresos de este periodo timbrados en otro mes
        totalOpsFacturadasOtroMes: totalOpsFacturadasOtroMes,
        xmlDeOtroPeriodo: huerfanosOtroPeriodo,             // CFDIs timbrados este mes de ventas de otro mes
        totalXmlDeOtroPeriodo: totalXmlDeOtroPeriodo
      },
      // XML sin dueño en el periodo pero cuyo monto coincide con una venta
      // pendiente (sin folio) de un mes colindante — no son sobrante, son
      // candidatos directos a vincular a esa operación específica.
      pendientesVecinos: pendientesVecinos,
      totalPendientesVecinos: totalPendientesVecinos,
      resumenFacturacion: {
        totalFacturadoXml: totalFacturadoXml, totalIngresos: totalIngresos,
        totalHuerfanos: totalHuerfanos, diferencia: totalFacturadoXml - totalIngresos,
        totalSobrantes: totalSobrantes, totalXmlDeOtroPeriodo: totalXmlDeOtroPeriodo,
        totalOpsFacturadasOtroMes: totalOpsFacturadasOtroMes,
        totalNotasCredito: totalNotasCredito, totalPendientesVecinos: totalPendientesVecinos
      }
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Búsqueda manual de XML en un rango de meses AMPLIO (más allá del ±1 mes
// automático) — para cuando el nombre/monto no coinciden con lo que ya está
// registrado en Ingresos (ej. la factura la timbraron con una razón social
// distinta al paciente, o el CFDI cayó fuera del padding por defecto). El
// usuario revisa la lista y decide a qué operación vincular cada XML — nunca
// se auto-asigna nada aquí, solo se descartan los que YA están vinculados a
// alguna operación (para no ofrecer dos veces el mismo CFDI).
function buscarXmlAmplio(fechaInicio, fechaFin, padMeses) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    var pad = parseInt(padMeses, 10);
    if (!pad || pad < 1) pad = 3;
    if (pad > 12) pad = 12; // límite razonable — más que esto tarda demasiado en Drive

    var months = _facMonthsInRangePadded(fechaInicio, fechaFin, pad);
    var duenoIdx = _facIndexDuenosBDIngresos();
    var results = [];
    var carpetas = [];

    months.forEach(function (ym) {
      var mesTag = (ym.mes < 10 ? '0' : '') + ym.mes + ' ' + FAC_MESES_ABR[ym.mes - 1];
      var folder = _facMonthFolder(ym.anio, ym.mes);
      carpetas.push({ path: ym.anio + '/' + mesTag, encontrada: !!folder });
      if (!folder) return;
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var name = file.getName();
        if (!/\.xml$/i.test(name)) continue;
        var xml;
        try { xml = file.getBlob().getDataAsString('UTF-8'); } catch (e) { continue; }
        var p = _facQuickParse(xml);
        if (p.tipo !== 'I') continue;
        var fileId = file.getId();
        var uuidUp = p.uuid ? p.uuid.toUpperCase() : '';
        var nf = p.folio ? _facNormFolio(p.folio) : '';
        var asignado = duenoIdx.porFileId[fileId] || (uuidUp && duenoIdx.porUuid[uuidUp]) || (nf && duenoIdx.porFolio[nf]);
        if (asignado) continue; // ya vinculado a alguna operación — no ofrecerlo de nuevo
        var fechaISO = p.fecha ? p.fecha.substring(0, 10) : '';
        results.push({
          folio: p.folio, serie: p.serie, fecha: fechaISO, mes: fechaISO.substring(0, 7),
          total: p.total, receptorNombre: p.receptorNombre, receptorRfc: p.receptorRfc, uuid: p.uuid,
          fileId: fileId, fileUrl: file.getUrl()
        });
      }
    });

    results.sort(function (a, b) { return a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0; }); // más reciente primero
    return { ok: true, rows: results, mesesEscaneados: months.length, carpetas: carpetas };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Índice de "dueños" de facturas en TODO BD_Ingresos (sin filtro de fecha):
// fileId del ArchivoURL, UUID y folio normalizado → {opId, fecha, paciente}.
// Sirve para saber si un XML que parece huérfano en este periodo en realidad
// pertenece a una operación de otro mes.
function _facIndexDuenosBDIngresos() {
  var out = { porFileId: {}, porUuid: {}, porFolio: {} };
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sheet) return out;
    var raw = sheet.getDataRange().getValues();
    var hdrs = (raw[0] || []).map(function (h) { return String(h).trim().toLowerCase(); });
    var iUuid = hdrs.indexOf('facturauuid');
    function dt(v) {
      if (!v) return '';
      if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
      return String(v).substring(0, 10);
    }
    for (var i = 1; i < raw.length; i++) {
      var r = raw[i];
      var op = String(r[0] || '').trim();
      if (!op) continue;
      var info = { opId: op, fecha: dt(r[2]), paciente: String(r[3] || '') };
      var url = String(r[22] || '').trim();
      if (url) { var m = url.match(/[-\w]{25,}/); if (m && !out.porFileId[m[0]]) out.porFileId[m[0]] = info; }
      var uuid = iUuid > -1 ? String(r[iUuid] || '').trim().toUpperCase() : '';
      if (uuid && !out.porUuid[uuid]) out.porUuid[uuid] = info;
      var folio = String(r[17] || '').trim();
      if (folio) { var nf = _facNormFolio(folio); if (nf && !out.porFolio[nf]) out.porFolio[nf] = info; }
    }
  } catch (e) {}
  return out;
}

// Agrega FacturaRFC y FacturaUUID a BD_Ingresos si no existen (migración segura, al final)
function migrateBDIngresosFacturaDetalle() {
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sh) return { ok: false, error: 'BD_Ingresos no encontrada' };
    var raw = sh.getDataRange().getValues();
    if (!raw.length) return { ok: false, error: 'Hoja vacía' };
    var existing = raw[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var toAdd = ['FacturaRFC', 'FacturaUUID'].filter(function (h) { return existing.indexOf(h.toLowerCase()) === -1; });
    if (toAdd.length) {
      var lastCol = sh.getLastColumn();
      sh.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
      sh.getRange(1, lastCol + 1, 1, toAdd.length).setFontWeight('bold').setBackground('#fce7f3');
    }
    return { ok: true, agregadas: toAdd };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function vincularXmlFactura(body) {
  try {
    var opId = String(body.opId || '').trim();
    var fileId = String(body.fileId || '').trim();
    if (!opId || !fileId) return { ok: false, error: 'opId y fileId requeridos' };
    var file = DriveApp.getFileById(fileId);
    var url = file.getUrl();
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sheet) return { ok: false, error: 'BD_Ingresos no encontrada' };
    migrateBDIngresosFacturaDetalle();
    var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var urlCol = BD_INGRESOS_HEADERS.indexOf('ArchivoURL') + 1;
    var razonCol = hdrs.indexOf('RazonSocial') + 1;
    var rfcCol = hdrs.indexOf('FacturaRFC') + 1;
    var uuidCol = hdrs.indexOf('FacturaUUID') + 1;
    var data = sheet.getDataRange().getValues();
    var updated = 0;
    for (var ri = 1; ri < data.length; ri++) {
      if (String(data[ri][0]) !== opId) continue;
      sheet.getRange(ri + 1, urlCol).setValue(url);
      // Si el frontend ya trae la razón social/RFC/UUID del XML (viene de la
      // conciliación, que ya lo leyó), se guarda de una vez — evita depender del
      // backfill para las vinculaciones nuevas.
      if (razonCol > 0 && body.razonSocial) sheet.getRange(ri + 1, razonCol).setValue(body.razonSocial);
      if (rfcCol > 0 && body.rfc) sheet.getRange(ri + 1, rfcCol).setValue(body.rfc);
      if (uuidCol > 0 && body.uuid) sheet.getRange(ri + 1, uuidCol).setValue(body.uuid);
      updated++;
    }
    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (e) {}
    logAudit(body.usuario || 'sistema', 'Ingreso', 'VincularXML', opId, 'ArchivoURL', '', url);
    return { ok: true, url: url, updated: updated };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Sube el XML que el usuario tiene a mano y vincula la operación en un solo
// paso, leyendo el Folio/Serie/UUID/Razón social/RFC directo del propio
// archivo — el usuario ya NO tiene que escribir el folio a mano (muchas
// facturas se timbran a "PÚBLICO EN GENERAL", sin nombre que dé pie a una
// sugerencia automática, así que subir el XML manualmente es la única forma
// de vincularlas).
function uploadYVincularXmlFactura(body) {
  try {
    var opId = String(body.opId || '').trim();
    var base64Data = body.base64;
    if (!opId || !base64Data) return { ok: false, error: 'opId y archivo requeridos' };
    if (!INGRESOS_FOLDER_FACTURAS) return { ok: false, error: 'Carpeta de Drive no configurada para facturas' };

    var fileName = String(body.fileName || (opId + '.xml'));
    var folder = DriveApp.getFolderById(INGRESOS_FOLDER_FACTURAS);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'text/xml', fileName);
    var file = folder.createFile(blob);
    file.setName(opId + '_' + fileName);
    var fileId = file.getId();
    var url = file.getUrl();

    var parsed = _facParseCfdiFull(fileId);
    if (!parsed || !parsed.ok) return { ok: false, error: 'El archivo se subió pero no se pudo leer como XML de factura: ' + (parsed ? parsed.error : 'desconocido') };
    if (!parsed.folio) return { ok: false, error: 'El XML no trae número de Folio — verifica que sea el CFDI correcto.' };

    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sheet) return { ok: false, error: 'BD_Ingresos no encontrada' };
    migrateBDIngresosFacturaDetalle();
    var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var urlCol = BD_INGRESOS_HEADERS.indexOf('ArchivoURL') + 1;
    var facturaCol = BD_INGRESOS_HEADERS.indexOf('Factura') + 1;
    var razonCol = hdrs.indexOf('RazonSocial') + 1;
    var rfcCol = hdrs.indexOf('FacturaRFC') + 1;
    var uuidCol = hdrs.indexOf('FacturaUUID') + 1;
    var data = sheet.getDataRange().getValues();
    var updated = 0;
    for (var ri = 1; ri < data.length; ri++) {
      if (String(data[ri][0]) !== opId) continue;
      sheet.getRange(ri + 1, urlCol).setValue(url);
      sheet.getRange(ri + 1, facturaCol).setValue(parsed.folio);
      if (razonCol > 0) sheet.getRange(ri + 1, razonCol).setValue(parsed.receptor.nombre || '');
      if (rfcCol > 0) sheet.getRange(ri + 1, rfcCol).setValue(parsed.receptor.rfc || '');
      if (uuidCol > 0) sheet.getRange(ri + 1, uuidCol).setValue(parsed.uuid || '');
      updated++;
    }
    if (!updated) return { ok: false, error: 'No se encontró la operación ' + opId + ' en BD_Ingresos' };
    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (e) {}
    logAudit(body.usuario || 'sistema', 'Ingreso', 'SubirVincularXML', opId, 'ArchivoURL', '', url);
    return {
      ok: true, url: url, updated: updated, folio: parsed.folio, serie: parsed.serie,
      razonSocial: parsed.receptor.nombre || '', rfc: parsed.receptor.rfc || '', uuid: parsed.uuid || '', total: parsed.total
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Aplica de un jalón TODAS las sugerencias sin ambigüedad que ya encontró
// reconciliarFacturasXml (un solo candidato, sea por nombre o por monto) —
// sin pedir ningún archivo: el XML ya está localizado en Drive desde el
// escaneo de la conciliación, así que solo hace falta guardar la referencia,
// igual que el botón "Usar y vincular" pero para todas a la vez. Nace de que
// pedirle al usuario que seleccione manualmente cuáles XML subir no sirve —
// los archivos se llaman por UUID, no traen folio/nombre/monto visible, así
// que no hay forma de saber cuáles ya están vinculados sin abrirlos uno por
// uno. Los casos ambiguos (varias operaciones con el mismo monto) NO se
// tocan aquí — siguen resolviéndose a mano en "Documentos faltantes".
function vincularAutomaticoLote(fechaInicio, fechaFin, usuario) {
  try {
    var rec = reconciliarFacturasXml(fechaInicio, fechaFin);
    if (!rec.ok) return rec;
    var candidatos = rec.sinFactura.filter(function (o) { return o.sugerenciaFolio; });
    var ambiguas = rec.sinFactura.filter(function (o) { return o.candidatosPorMonto && o.candidatosPorMonto.length; }).length;
    var sinNada = rec.sinFactura.length - candidatos.length - ambiguas;
    // faltaDocumento = el folio capturado YA coincide con un XML localizado —
    // vínculo sin ambigüedad, solo falta guardar la URL. Entra al lote también.
    var porFolio = rec.faltaDocumento || [];
    if (!candidatos.length && !porFolio.length) return { ok: true, vinculadas: [], ambiguas: ambiguas, sinNada: sinNada };

    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sheet = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sheet) return { ok: false, error: 'BD_Ingresos no encontrada' };
    migrateBDIngresosFacturaDetalle();
    var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var urlCol = BD_INGRESOS_HEADERS.indexOf('ArchivoURL') + 1;
    var facturaCol = BD_INGRESOS_HEADERS.indexOf('Factura') + 1;
    var razonCol = hdrs.indexOf('RazonSocial') + 1;
    var rfcCol = hdrs.indexOf('FacturaRFC') + 1;
    var uuidCol = hdrs.indexOf('FacturaUUID') + 1;
    var data = sheet.getDataRange().getValues();
    var rowsByOpId = {};
    for (var ri = 1; ri < data.length; ri++) {
      var opId0 = String(data[ri][0]);
      if (!rowsByOpId[opId0]) rowsByOpId[opId0] = [];
      rowsByOpId[opId0].push(ri + 1);
    }

    var vinculadas = [];
    function escribir(opId, fileId, folio, razonSocial, rfc, uuid, total, paciente, tipo) {
      var file;
      try { file = DriveApp.getFileById(fileId); } catch (e) { return; }
      var url = file.getUrl();
      var filas = rowsByOpId[opId] || [];
      filas.forEach(function (rowNum) {
        sheet.getRange(rowNum, urlCol).setValue(url);
        if (folio) sheet.getRange(rowNum, facturaCol).setValue(folio);
        if (razonCol > 0 && razonSocial) sheet.getRange(rowNum, razonCol).setValue(razonSocial);
        if (rfcCol > 0 && rfc) sheet.getRange(rowNum, rfcCol).setValue(rfc);
        if (uuidCol > 0 && uuid) sheet.getRange(rowNum, uuidCol).setValue(uuid);
      });
      if (filas.length) vinculadas.push({ opId: opId, paciente: paciente, folio: folio, razonSocial: razonSocial || '', rfc: rfc || '', total: total, tipo: tipo });
    }

    porFolio.forEach(function (fd) {
      escribir(fd.id, fd.xmlFileId, fd.factura, fd.xmlRazonSocial, fd.xmlRfc, fd.xmlUuid, fd.total, fd.paciente, 'folio');
    });
    candidatos.forEach(function (op) {
      escribir(op.id, op.sugerenciaFileId, op.sugerenciaFolio, op.sugerenciaRazonSocial, op.sugerenciaRfc, op.sugerenciaUuid, op.total, op.paciente, op.sugerenciaTipo);
    });

    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (e) {}
    logAudit(usuario || 'sistema', 'Ingreso', 'VincularAutomaticoLote', '', fechaInicio + ' a ' + fechaFin, '', vinculadas.length + ' vinculadas automáticamente');
    return { ok: true, vinculadas: vinculadas, ambiguas: ambiguas, sinNada: sinNada };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Completa Razón Social/RFC/UUID de operaciones YA vinculadas (ArchivoURL) usando
// el mismo escaneo de XML de la conciliación — cruza por fileId (extraído de la
// URL de Drive guardada) contra el índice de XML del periodo. Pensado para
// operaciones vinculadas antes de que existiera este detalle, o vinculadas
// manualmente sin pasar por vincularXmlFactura. Siempre sincroniza con el XML
// (fuente de verdad de a quién se facturó realmente), aunque ya hubiera algo
// capturado a mano.
function backfillRazonSocialDesdeXml(fechaInicio, fechaFin, usuario) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    migrateBDIngresosFacturaDetalle();
    var idx = _facBuildXmlIndex(fechaInicio, fechaFin);
    var byFileId = {};
    idx.all.forEach(function (x) { byFileId[x.fileId] = x; });

    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sh) return { ok: false, error: 'BD_Ingresos no encontrada' };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function (h) { return String(h).trim(); });
    var idxArchivo = hdrs.indexOf('ArchivoURL');
    var idxFecha = hdrs.indexOf('Fecha');
    var idxRazon = hdrs.indexOf('RazonSocial');
    var idxRFC = hdrs.indexOf('FacturaRFC');
    var idxUUID = hdrs.indexOf('FacturaUUID');
    if (idxArchivo < 0) return { ok: false, error: 'Columna ArchivoURL no encontrada' };

    function dt(v) {
      if (!v) return '';
      if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
      return String(v).substring(0, 10);
    }

    var actualizadas = 0, sinCoincidencia = 0;
    for (var i = 1; i < data.length; i++) {
      var url = String(data[i][idxArchivo] || '').trim();
      if (!url) continue;
      var fechaStr = dt(data[i][idxFecha]);
      if (fechaStr < fechaInicio || fechaStr > fechaFin) continue;
      var m = url.match(/[-\w]{25,}/);
      if (!m) continue;
      var x = byFileId[m[0]];
      if (!x) { sinCoincidencia++; continue; }
      if (idxRazon > -1) sh.getRange(i + 1, idxRazon + 1).setValue(x.receptorNombre || '');
      if (idxRFC > -1) sh.getRange(i + 1, idxRFC + 1).setValue(x.receptorRfc || '');
      if (idxUUID > -1) sh.getRange(i + 1, idxUUID + 1).setValue(x.uuid || '');
      actualizadas++;
    }
    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (e) {}
    logAudit(usuario || 'sistema', 'Ingreso', 'BackfillRazonSocial', '', fechaInicio + ' a ' + fechaFin, '', actualizadas + ' filas actualizadas');
    return { ok: true, actualizadas: actualizadas, sinCoincidencia: sinCoincidencia };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Parseo completo (XmlService) para el reporte — incluye impuestos por concepto
function _facParseCfdiFull(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var content = file.getBlob().getDataAsString('UTF-8');
    return _facParseCfdiFullFromContent(content);
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Igual que _facParseCfdiFull pero a partir del texto del XML directamente
// (sin necesidad de que el archivo ya exista en Drive) — usado por la subida
// masiva, que valida/empareja el XML antes de decidir si vale la pena guardarlo.
function _facParseCfdiFullFromContent(content) {
  try {
    if (content.charCodeAt(0) === 0xFEFF) content = content.substring(1);
    content = content.replace(/^\s+/, '');
    var doc = XmlService.parse(content);
    var root = doc.getRootElement();
    var ns = root.getNamespace();
    var attr = function (el, name) { var a = el.getAttribute(name); return a ? a.getValue() : ''; };
    var child = function (el, name) { return el.getChild(name, ns); };
    var emisorEl = child(root, 'Emisor');
    var receptorEl = child(root, 'Receptor');
    var conceptos = [];
    var conceptosEl = child(root, 'Conceptos');
    if (conceptosEl) {
      var cEls = conceptosEl.getChildren('Concepto', ns);
      for (var i = 0; i < cEls.length; i++) {
        var c = cEls[i];
        var ivaPct = 'EXENTO', ivaMonto = 0;
        var impEl = c.getChild('Impuestos', ns);
        if (impEl) {
          var traslEl = impEl.getChild('Traslados', ns);
          if (traslEl) {
            var ts = traslEl.getChildren('Traslado', ns);
            if (ts.length) {
              var factor = attr(ts[0], 'TipoFactor');
              if (factor !== 'Exento') {
                ivaPct = parseFloat(attr(ts[0], 'TasaOCuota')) || 0;
                ivaMonto = parseFloat(attr(ts[0], 'Importe')) || 0;
              }
            }
          }
        }
        conceptos.push({
          claveProdServ: attr(c, 'ClaveProdServ'), noIdentificacion: attr(c, 'NoIdentificacion'),
          cantidad: parseFloat(attr(c, 'Cantidad')) || 1, descripcion: attr(c, 'Descripcion'),
          valorUnitario: parseFloat(attr(c, 'ValorUnitario')) || 0, importe: parseFloat(attr(c, 'Importe')) || 0,
          descuento: parseFloat(attr(c, 'Descuento')) || 0,
          ivaPct: ivaPct, ivaMonto: ivaMonto
        });
      }
    }
    // TimbreFiscalDigital vive dentro de Complemento pero en el namespace tfd
    // (no cfdi), así que no se puede pedir con getChild(nombre, ns) — se busca
    // por nombre local entre todos los hijos, sin filtrar por namespace.
    var uuid = '';
    var complementoEl = child(root, 'Complemento');
    if (complementoEl) {
      var compChildren = complementoEl.getChildren();
      for (var ki = 0; ki < compChildren.length; ki++) {
        if (compChildren[ki].getName() === 'TimbreFiscalDigital') { uuid = attr(compChildren[ki], 'UUID'); break; }
      }
    }
    return {
      ok: true, serie: attr(root, 'Serie'), folio: attr(root, 'Folio'), fecha: attr(root, 'Fecha'), uuid: uuid,
      subTotal: parseFloat(attr(root, 'SubTotal')) || 0, descuento: parseFloat(attr(root, 'Descuento')) || 0,
      total: parseFloat(attr(root, 'Total')) || 0,
      moneda: attr(root, 'Moneda') || 'MXN', tipoCambio: parseFloat(attr(root, 'TipoCambio')) || 1,
      formaPago: attr(root, 'FormaPago'), metodoPago: attr(root, 'MetodoPago'),
      totalImpuestosTrasladados: (function () { var im = child(root, 'Impuestos'); return im ? (parseFloat(attr(im, 'TotalImpuestosTrasladados')) || 0) : 0; })(),
      emisor: { rfc: emisorEl ? attr(emisorEl, 'Rfc') : '', nombre: emisorEl ? attr(emisorEl, 'Nombre') : '' },
      receptor: { rfc: receptorEl ? attr(receptorEl, 'Rfc') : '', nombre: receptorEl ? attr(receptorEl, 'Nombre') : '', usoCfdi: receptorEl ? attr(receptorEl, 'UsoCFDI') : '' },
      conceptos: conceptos
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Índice paciente (nombre normalizado) -> {razonSocial, rfc} desde Registro de Pacientes
function _pacFiscalIndex() {
  var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
  var sh = ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var hdrs = (data[0] || []).map(function (h) { return String(h).trim(); });
  var idxRS = _pacColIdx(hdrs, 'Razon Social'), idxRFC = _pacColIdx(hdrs, 'RFC');
  var byNorm = {};
  for (var i = 1; i < data.length; i++) {
    var nombre = String(data[i][1] || '').trim();
    if (!nombre) continue;
    var rfc = idxRFC > -1 ? String(data[i][idxRFC] || '').trim() : '';
    var rs = idxRS > -1 ? String(data[i][idxRS] || '').trim() : '';
    if (!rfc && !rs) continue;
    byNorm[_pacNormNombre(nombre)] = { razonSocial: rs, rfc: rfc };
  }
  return byNorm;
}

function generarReporteContaDigital(fechaInicio, fechaFin, usuario) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    var ops = _facReadOpsInRange(fechaInicio, fechaFin);
    var idx = _facBuildXmlIndex(fechaInicio, fechaFin);
    var byFileId = {};
    idx.all.forEach(function (x) { byFileId[x.fileId] = x; });

    // El dato con el que SIEMPRE se puede llegar al XML real es ArchivoURL (lo
    // que llena "Vincular"/"Analizar periodo" en Facturación) — no el campo
    // Factura# capturado a mano, que muchas veces no coincide con el folio del
    // XML (ej. facturas vinculadas por sugerencia de nombre, no por folio).
    var facturables = [];
    ops.forEach(function (op) {
      if (!op.archivoURL) return;
      var m = op.archivoURL.match(/[-\w]{25,}/);
      if (!m) return;
      var x = byFileId[m[0]];
      if (!x) return;
      facturables.push({ op: op, fileId: m[0] });
    });
    if (!facturables.length) return { ok: false, error: 'No hay operaciones con XML vinculado en el periodo seleccionado. Usa "Analizar periodo actual" en Facturación para vincular primero.' };

    var pacFiscal = _pacFiscalIndex();

    var invoices = [];
    facturables.forEach(function (item) {
      var full = _facParseCfdiFull(item.fileId);
      if (!full || !full.ok || !full.conceptos.length) return;
      var pf = pacFiscal[_pacNormNombre(item.op.paciente)];
      full._paciente = item.op.paciente;
      full._opId = item.op.id;
      // Regla: datos fiscales del Registro de Pacientes tienen prioridad; si el
      // paciente no tiene RFC/Razón Social capturados, se usa público en general.
      full._rfcFinal = (pf && pf.rfc) ? pf.rfc : 'XAXX010101000';
      full._razonSocialFinal = (pf && pf.razonSocial) ? pf.razonSocial : 'PUBLICO EN GENERAL';
      invoices.push(full);
    });
    if (!invoices.length) return { ok: false, error: 'No se pudo leer ningún XML vinculado en el periodo' };

    // La plantilla real de ContaDigital siempre trae el mismo número de
    // columnas (17 + 10 bloques de concepto de 14 + 11 finales = 168) — se
    // respeta ese ancho fijo aunque una factura tenga menos o (recortando) más
    // líneas, para que el archivo sea consistente factura a factura.
    var headers = FAC_MASIVA_HEADERS_INV.slice();
    for (var c = 0; c < FAC_MASIVA_CONCEPTO_BLOQUES; c++) headers = headers.concat(FAC_MASIVA_HEADERS_CONCEPTO);
    headers = headers.concat(FAC_MASIVA_HEADERS_TRAIL);

    var rows = [headers];
    invoices.forEach(function (f) {
      var metodoLabel = FAC_METODO_PAGO_MAP[f.formaPago] || f.formaPago || '';
      var descuentoFactura = f.descuento || 0;
      var monedaLabel = FAC_MONEDA_MAP[f.moneda] || f.moneda || 'PESOS';
      // IMPORTE = bruto antes de descuento (= SubTotal del CFDI, que ya de por sí es
      // la suma de Importe por concepto sin descontar) · SUBTOTAL (columna propia de
      // ContaDigital) = neto antes de impuestos · TOTAL FACTURA = el Total real del CFDI.
      var row = [
        FAC_MASIVA_SERIE, f.fecha ? f.fecha.substring(0, 10) : '', '', f._rfcFinal,
        monedaLabel, (f.moneda && f.moneda !== 'MXN') ? (f.tipoCambio || 1) : '', metodoLabel, '', f._opId,
        f.subTotal, descuentoFactura, f.subTotal - descuentoFactura, f.totalImpuestosTrasladados || 0, 0, 0, 0, f.total
      ];
      var nConceptos = Math.min(f.conceptos.length, FAC_MASIVA_CONCEPTO_BLOQUES);
      for (var ci = 0; ci < nConceptos; ci++) {
        var con = f.conceptos[ci];
        var ivaPctVal = (con.ivaPct === 'EXENTO') ? 0 : con.ivaPct;
        row.push(
          con.noIdentificacion || con.claveProdServ || '', con.descripcion || '', con.cantidad, con.valorUnitario, con.descuento || 0,
          ivaPctVal, con.ivaMonto, '', 0, 'No aplica', 0, '', 0, ''
        );
      }
      for (var mi = nConceptos; mi < FAC_MASIVA_CONCEPTO_BLOQUES; mi++) row = row.concat(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      row.push(f.receptor.usoCfdi || '', '', f.metodoPago || '', 'Factura', '', '', '', '', '', '', '');
      rows.push(row);
    });

    // Se escribe el .xlsx directamente en formato OOXML (sin pasar por Sheets ni
    // por conversiones de Drive) para no depender de getAs()/UrlFetchApp, que han
    // fallado en este proyecto, ni de ningún permiso adicional. Solo usa
    // Utilities/DriveApp, que ya están autorizados en este proyecto.
    var xlsxBlob = _buildXlsxBlob(rows, 'Hoja1', 'ContaDigital_Masiva_' + fechaInicio + '_a_' + fechaFin + '.xlsx');
    var folder = DriveApp.getFolderById(FAC_MASIVA_FOLDER_ID);
    var xlsxFile = folder.createFile(xlsxBlob);

    var detalle = invoices.map(function (f) {
      return {
        opId: f._opId, paciente: f._paciente, razonSocial: f._razonSocialFinal,
        rfc: f._rfcFinal, folio: f.folio, serie: f.serie,
        fecha: f.fecha ? f.fecha.substring(0, 10) : '', total: f.total,
        subTotal: f.subTotal, descuento: f.descuento || 0
      };
    });
    var totalGeneral = detalle.reduce(function (s, d) { return s + (d.total || 0); }, 0);

    logAudit(usuario || 'sistema', 'Facturacion', 'ReporteContaDigital', '', fechaInicio + ' a ' + fechaFin, '', invoices.length + ' facturas');
    return {
      ok: true, url: xlsxFile.getUrl(), numFacturas: invoices.length,
      numOperacionesTotal: ops.length, numSinXml: ops.length - facturables.length,
      detalle: detalle, totalGeneral: totalGeneral
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Reporte ContaDigital de lo que TODAVÍA NO se ha facturado — no lee XML (no
// existe) sino la venta real capturada en BD_Ingresos, para poder subir el
// archivo a ContaDigital y generar ahí las facturas faltantes.
// Regla de seguridad: solo entran operaciones sin ArchivoURL, sin folio que
// haga match con un XML del periodo, Y SIN sugerencia de nombre encontrada por
// reconciliarFacturasXml — si hay una sugerencia, lo más probable es que la
// factura YA exista y solo falte vincularla, así que se excluye para no
// arriesgar una factura duplicada.
function generarReporteContaDigitalPendientes(fechaInicio, fechaFin, usuario) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    var rec = reconciliarFacturasXml(fechaInicio, fechaFin);
    if (!rec.ok) return rec;
    var pendientesIds = rec.sinFactura.filter(function (o) { return !o.sugerenciaFolio; }).map(function (o) { return o.id; });
    if (!pendientesIds.length) return { ok: false, error: 'No hay operaciones pendientes por facturar en el periodo seleccionado (todas ya tienen factura o una sugerencia de XML por confirmar en la sección de arriba).' };

    var opsMap = _facReadOpsPendientesDetalle(fechaInicio, fechaFin, pendientesIds);
    var pacFiscal = _pacFiscalIndex();

    // Porción pagada con Nota de Crédito por OP: es descuento — la factura se
    // timbra por lo realmente cobrado, así que se resta del total y se suma
    // al campo DESCUENTO de la plantilla.
    var ncPorOp = {};
    rec.sinFactura.forEach(function (o) { if (o.notaCredito > 0) ncPorOp[o.id] = o.notaCredito; });

    var headers = FAC_MASIVA_HEADERS_INV.slice();
    for (var c = 0; c < FAC_MASIVA_CONCEPTO_BLOQUES; c++) headers = headers.concat(FAC_MASIVA_HEADERS_CONCEPTO);
    headers = headers.concat(FAC_MASIVA_HEADERS_TRAIL);

    var rows = [headers];
    var detalle = [];
    pendientesIds.forEach(function (id) {
      var op = opsMap[id];
      if (!op || !op.lineas.length) return;
      var pf = pacFiscal[_pacNormNombre(op.paciente)];
      var rfcFinal = (pf && pf.rfc) ? pf.rfc : 'XAXX010101000';
      var razonFinal = (pf && pf.razonSocial) ? pf.razonSocial : 'PUBLICO EN GENERAL';
      var nc = ncPorOp[op.id] || 0;
      var totalCobrado = Math.max(0, op.total - nc); // la NC es descuento — se factura lo cobrado
      var descuentoOp = Math.max(0, op.montoLista - totalCobrado);
      // Sin CFDI del que leer impuestos: se asume Exento (servicios médicos, el
      // caso normal en Hestia) — revisar antes de subir si alguna línea no lo es.
      var row = [
        FAC_MASIVA_SERIE, op.fecha, op.sucursal || '', rfcFinal,
        'PESOS', '', op.formaPago || '', '', op.id,
        op.montoLista, descuentoOp, totalCobrado, 0, 0, 0, 0, totalCobrado
      ];
      var nLineas = Math.min(op.lineas.length, FAC_MASIVA_CONCEPTO_BLOQUES);
      for (var li = 0; li < nLineas; li++) {
        var ln = op.lineas[li];
        var descPesos = ln.pvp * ln.cantidad * (ln.descPct / 100);
        row.push('', ln.producto || '', ln.cantidad, ln.pvp, descPesos, 0, 0, '', 0, 'No aplica', 0, '', 0, '');
      }
      for (var mi = nLineas; mi < FAC_MASIVA_CONCEPTO_BLOQUES; mi++) row = row.concat(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      row.push('', '', '', 'Factura', '', '', '', '', '', '', '');
      rows.push(row);
      detalle.push({ opId: op.id, paciente: op.paciente, razonSocial: razonFinal, rfc: rfcFinal, fecha: op.fecha, total: totalCobrado, notaCredito: nc });
    });
    if (rows.length < 2) return { ok: false, error: 'No se pudo armar el detalle de las operaciones pendientes (sin líneas de venta en el rango).' };

    var xlsxBlob = _buildXlsxBlob(rows, 'Hoja1', 'ContaDigital_PorFacturar_' + fechaInicio + '_a_' + fechaFin + '.xlsx');
    var folder = DriveApp.getFolderById(FAC_MASIVA_FOLDER_ID);
    var xlsxFile = folder.createFile(xlsxBlob);

    var totalGeneral = detalle.reduce(function (s, d) { return s + (d.total || 0); }, 0);
    logAudit(usuario || 'sistema', 'Facturacion', 'ReporteContaDigitalPendientes', '', fechaInicio + ' a ' + fechaFin, '', detalle.length + ' operaciones pendientes');
    return {
      ok: true, url: xlsxFile.getUrl(), numOperaciones: detalle.length,
      detalle: detalle, totalGeneral: totalGeneral
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ============================================================
// DATOS FISCALES DE PACIENTES — cruce histórico XML → Ingresos → Pacientes
// ============================================================

var PAC_FISCAL_HEADERS = ['Razon Social', 'RFC', 'Codigo Postal', 'Uso CFDI', 'Regimen Fiscal', 'Forma de Pago Habitual'];

function _pacColIdx(headers, name) {
  var nl = String(name).trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === nl) return i;
  }
  return -1;
}

function _pacNormNombre(s) {
  var out = String(s || '').trim().toLowerCase().normalize('NFD');
  var stripped = '';
  for (var i = 0; i < out.length; i++) {
    var code = out.charCodeAt(i);
    if (code < 0x0300 || code > 0x036f) stripped += out[i];
  }
  return stripped.replace(/\s+/g, ' ');
}

// Agrega las 6 columnas fiscales a Pacientes si no existen (migración segura, al final)
function migratePacientesFiscales() {
  try {
    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var raw = sh.getDataRange().getValues();
    if (!raw.length) return { ok: false, error: 'Hoja Pacientes vacía' };
    var existing = raw[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var toAdd = PAC_FISCAL_HEADERS.filter(function (h) { return existing.indexOf(h.toLowerCase()) === -1; });
    if (toAdd.length) {
      var lastCol = sh.getLastColumn();
      sh.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
      sh.getRange(1, lastCol + 1, 1, toAdd.length).setFontWeight('bold').setBackground('#fce7f3');
    }
    return { ok: true, agregadas: toAdd };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Agrega la columna RazonSocial a BD_Ingresos si no existe (migración segura, al final)
function migrateBDIngresosRazonSocial() {
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sh) return { ok: false, error: 'BD_Ingresos no encontrada' };
    var raw = sh.getDataRange().getValues();
    if (!raw.length) return { ok: false, error: 'Hoja vacía' };
    var existing = raw[0].map(function (h) { return String(h).trim().toLowerCase(); });
    if (existing.indexOf('razonsocial') === -1 && existing.indexOf('razon social') === -1) {
      var lastCol = sh.getLastColumn();
      sh.getRange(1, lastCol + 1).setValue('RazonSocial');
      sh.getRange(1, lastCol + 1).setFontWeight('bold').setBackground('#fce7f3');
      return { ok: true, agregada: true };
    }
    return { ok: true, agregada: false };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Analiza (SOLO LECTURA, no escribe nada) el historial de XML del rango contra Ingresos
// y contra el Registro de Pacientes, para proponer Razón Social/RFC/CP/Uso CFDI/Régimen/
// Forma de pago por paciente. Regla: la factura más reciente gana; si hay razones sociales
// distintas en el historial, se marca como conflicto y se arma una nota con el detalle.
function analizarDatosFiscalesPacientes(fechaInicio, fechaFin) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    migratePacientesFiscales();
    migrateBDIngresosRazonSocial();

    var idx = _facBuildXmlIndex(fechaInicio, fechaFin);
    var ops = _facReadOpsInRange(fechaInicio, fechaFin);

    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    var pacientes = [];
    for (var i = 1; i < data.length; i++) {
      var nombre = String(data[i][1] || '').trim();
      if (!nombre) continue;
      pacientes.push({ id: String(data[i][0] || '').trim(), nombre: nombre, norm: _pacNormNombre(nombre) });
    }
    var pacByNorm = {};
    pacientes.forEach(function (p) { pacByNorm[p.norm] = p; });

    var hallazgosPorPaciente = {};
    var sinMatchPaciente = [];
    var opsConFactura = 0;

    ops.forEach(function (op) {
      if (!op.factura) return;
      opsConFactura++;
      var x = _facXmlPorOp(idx, op);
      if (!x) return;
      var opNorm = _pacNormNombre(op.paciente);
      var pMatch = pacByNorm[opNorm];
      var matchType = 'exacto';
      if (!pMatch && opNorm) {
        for (var pi = 0; pi < pacientes.length; pi++) {
          var pn = pacientes[pi].norm;
          if (!pn) continue;
          if (pn.indexOf(opNorm) > -1 || opNorm.indexOf(pn) > -1) { pMatch = pacientes[pi]; matchType = 'parcial'; break; }
        }
      }
      var hallazgo = {
        opId: op.id, folio: op.factura, fecha: x.fecha, opTotal: op.total, xmlTotal: x.total,
        razonSocial: x.receptorNombre, rfc: x.receptorRfc, codigoPostal: x.receptorCP,
        usoCfdi: x.receptorUsoCfdi, regimenFiscal: x.receptorRegimen,
        formaPago: FAC_METODO_PAGO_MAP[x.formaPago] || x.formaPago || '',
        matchType: matchType, pacienteNombreOp: op.paciente
      };
      if (pMatch) {
        if (!hallazgosPorPaciente[pMatch.id]) hallazgosPorPaciente[pMatch.id] = { pacienteId: pMatch.id, pacienteNombre: pMatch.nombre, hallazgos: [] };
        hallazgosPorPaciente[pMatch.id].hallazgos.push(hallazgo);
      } else {
        sinMatchPaciente.push(hallazgo);
      }
    });

    var resultado = [];
    for (var pid in hallazgosPorPaciente) {
      var grupo = hallazgosPorPaciente[pid];
      grupo.hallazgos.sort(function (a, b) { return a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : 0; });
      var top = grupo.hallazgos[0];
      var distintos = {};
      grupo.hallazgos.forEach(function (h) { distintos[(h.razonSocial || '') + '|' + (h.rfc || '')] = true; });
      var conflicto = Object.keys(distintos).length > 1;
      var confianza = grupo.hallazgos.some(function (h) { return h.matchType === 'parcial'; }) ? 'parcial' : 'exacto';
      var notaHistorial = '';
      if (conflicto) {
        notaHistorial = 'Facturación histórica: ' + grupo.hallazgos.map(function (h) {
          return h.fecha + ' — ' + (h.razonSocial || 's/nombre') + ' (' + (h.rfc || 's/RFC') + ') — Folio ' + h.folio;
        }).join(' | ');
      }
      resultado.push({
        pacienteId: pid, pacienteNombre: grupo.pacienteNombre,
        razonSocial: top.razonSocial, rfc: top.rfc, codigoPostal: top.codigoPostal,
        usoCfdi: top.usoCfdi, regimenFiscal: top.regimenFiscal, formaPagoHabitual: top.formaPago,
        folioReferencia: top.folio, fechaReferencia: top.fecha,
        conflicto: conflicto, confianza: confianza, numFacturas: grupo.hallazgos.length,
        notaHistorial: notaHistorial
      });
    }
    resultado.sort(function (a, b) { return a.pacienteNombre < b.pacienteNombre ? -1 : 1; });

    return {
      ok: true, resultado: resultado, sinMatch: sinMatchPaciente,
      totalXmlEnRango: idx.all.filter(function (x) { return x.enPeriodo; }).length, totalOpsConFactura: opsConFactura,
      totalConMatch: resultado.length, totalSinMatch: sinMatchPaciente.length,
      totalConConflicto: resultado.filter(function (r) { return r.conflicto; }).length
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Escribe en Pacientes los datos fiscales confirmados por el usuario.
// body.confirmaciones: [{pacienteId, razonSocial, rfc, codigoPostal, usoCfdi, regimenFiscal, formaPagoHabitual, notaHistorial}]
function aplicarDatosFiscalesPacientes(body) {
  try {
    var confirmaciones = body.confirmaciones || [];
    if (!confirmaciones.length) return { ok: false, error: 'Sin confirmaciones' };
    migratePacientesFiscales();

    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function (h) { return String(h).trim(); });
    var idxRS = _pacColIdx(hdrs, 'Razon Social'), idxRFC = _pacColIdx(hdrs, 'RFC'),
        idxCP = _pacColIdx(hdrs, 'Codigo Postal'), idxUso = _pacColIdx(hdrs, 'Uso CFDI'),
        idxReg = _pacColIdx(hdrs, 'Regimen Fiscal'), idxFP = _pacColIdx(hdrs, 'Forma de Pago Habitual'),
        idxObs = _pacColIdx(hdrs, 'Observaciones / Notas');

    var rowByPacienteId = {};
    for (var i = 1; i < data.length; i++) rowByPacienteId[String(data[i][0] || '').trim()] = i + 1;

    var actualizados = 0;
    confirmaciones.forEach(function (c) {
      var rowNum = rowByPacienteId[c.pacienteId];
      if (!rowNum) return;
      if (idxRS > -1) sh.getRange(rowNum, idxRS + 1).setValue(c.razonSocial || '');
      if (idxRFC > -1) sh.getRange(rowNum, idxRFC + 1).setValue(c.rfc || '');
      if (idxCP > -1) sh.getRange(rowNum, idxCP + 1).setValue(c.codigoPostal || '');
      if (idxUso > -1) sh.getRange(rowNum, idxUso + 1).setValue(c.usoCfdi || '');
      if (idxReg > -1) sh.getRange(rowNum, idxReg + 1).setValue(c.regimenFiscal || '');
      if (idxFP > -1) sh.getRange(rowNum, idxFP + 1).setValue(c.formaPagoHabitual || '');
      if (idxObs > -1 && c.notaHistorial) {
        var actual = String(data[rowNum - 1][idxObs] || '');
        var nueva = actual ? (actual + ' | ' + c.notaHistorial) : c.notaHistorial;
        sh.getRange(rowNum, idxObs + 1).setValue(nueva);
      }
      actualizados++;
    });
    logAudit(body.usuario || 'sistema', 'Pacientes', 'SyncFiscal', '', '', '', actualizados + ' pacientes actualizados');
    return { ok: true, actualizados: actualizados };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ============================================================
// ANÁLISIS FISCAL DE DESCUENTOS
// ============================================================
// El CFDI 4.0 separa Importe (precio de lista, sin descuento) de Descuento
// (el descuento que el propio comprobante declara). El SAT prellena la
// declaración con el Importe/SubTotal bruto — el descuento hay que
// capturarlo aparte, a mano, en "Descuentos y devoluciones" para bajar la
// base gravable. El problema: algunas ventas se facturaron con el precio
// YA rebajado (sin usar el campo Descuento del CFDI), así que ese
// descuento nunca aparece declarable por separado — ya está absorbido en
// un Importe más bajo. Esta función distingue ambos casos comparando el
// % de descuento capturado en la venta (Ingresos) contra lo que el XML
// realmente declara.
function analizarDescuentosFiscales(fechaInicio, fechaFin) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    var ops = _facReadOpsConDetalleVenta(fechaInicio, fechaFin);
    var idx = _facBuildXmlIndex(fechaInicio, fechaFin);
    var byFileId = {};
    idx.all.forEach(function (x) { byFileId[x.fileId] = x; });

    var detalle = [];
    var totales = {
      montoFiscal: 0, descuentoCFDI: 0, montoHorneado: 0, totalFacturado: 0,
      sinDescuento: 0, reflejadoCfdi: 0, horneadoEnPrecio: 0, revisar: 0
    };

    ops.forEach(function (op) {
      if (!op.archivoURL) return;
      var m = op.archivoURL.match(/[-\w]{25,}/);
      if (!m) return;
      var x = byFileId[m[0]];
      if (!x) return; // factura vinculada pero el XML cayó fuera del rango escaneado

      var clasificacion;
      if (op.tieneDescuentoVenta && x.descuento > 0) clasificacion = 'reflejado_cfdi';
      else if (op.tieneDescuentoVenta && x.descuento === 0) clasificacion = 'horneado_en_precio';
      else if (!op.tieneDescuentoVenta && x.descuento > 0) clasificacion = 'revisar';
      else clasificacion = 'sin_descuento';

      var montoHorneado = clasificacion === 'horneado_en_precio' ? Math.max(0, op.montoLista - op.total) : 0;

      detalle.push({
        opId: op.id, fecha: op.fecha, paciente: op.paciente, factura: op.factura,
        opTotal: op.total, montoLista: op.montoLista, descuentoVentaPct: op.descuentoVentaPct,
        xmlSubTotal: x.subTotal, xmlDescuento: x.descuento, xmlTotal: x.total,
        clasificacion: clasificacion, montoHorneado: montoHorneado
      });

      totales.montoFiscal += x.subTotal;
      totales.descuentoCFDI += x.descuento;
      totales.montoHorneado += montoHorneado;
      totales.totalFacturado += x.total;
      if (clasificacion === 'sin_descuento') totales.sinDescuento++;
      else if (clasificacion === 'reflejado_cfdi') totales.reflejadoCfdi++;
      else if (clasificacion === 'horneado_en_precio') totales.horneadoEnPrecio++;
      else totales.revisar++;
    });

    detalle.sort(function (a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });

    return { ok: true, detalle: detalle, totales: totales };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
