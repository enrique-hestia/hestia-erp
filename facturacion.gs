// ============================================================
// FACTURACIÓN — Conciliación XML (Drive) vs Ingresos + Reporte ContaDigital
// ============================================================
// Estructura de carpetas en Drive (fija, ver memoria del proyecto):
// .../onefactureXMLs/HCL2307051Y6/emitidos/{año}/{MM MES}/  (archivos UUID.xml)

var FAC_MESES_ABR = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];

var FAC_MASIVA_HEADERS_INV = ['SERIE','FECHA','SUCURSAL','RFC RECEPTOR','MONEDA','T.C.','METODO DE PAGO','CUENTA BANCARIA','REFERENCIA','IMPORTE','DESCUENTO','SUBTOTAL','IVA TOTAL','IEPS TOTAL','RETENCION IVA','RETENCION ISR','TOTAL FACTURA'];
var FAC_MASIVA_HEADERS_CONCEPTO = ['CLAVE PRODUCTO','DESCRIPCION ADICIONAL','CANTIDAD','PRECIO','DESCUENTO $','IVA %','IVA $','IEPS %','IEPS $','RET IVA%','RET IVA $','RET ISR %','RET ISR $','PREDIAL'];

var FAC_METODO_PAGO_MAP = {
  '01':'Efectivo','02':'Cheque nominativo','03':'Transferencia electrónica de fondos',
  '04':'Tarjeta de crédito','05':'Monedero electrónico','06':'Dinero electrónico',
  '08':'Vales de despensa','28':'Tarjeta de débito','29':'Tarjeta de servicios',
  '31':'Tarjeta de crédito','99':'Tarjeta de crédito'
};

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

function _facQuickParse(xml) {
  var headEnd = xml.indexOf('<cfdi:Emisor');
  var head = headEnd > -1 ? xml.substring(0, headEnd) : xml;
  function a(name, src) { var m = (src || head).match(new RegExp(name + '="([^"]*)"')); return m ? m[1] : ''; }
  var recMatch = xml.match(/<cfdi:Receptor\s+([^>]*)\/>/);
  var rec = recMatch ? recMatch[1] : '';
  var uuidM = xml.match(/UUID="([^"]*)"/);
  return {
    folio: a('Folio'), serie: a('Serie'), tipo: a('TipoDeComprobante'),
    total: parseFloat(a('Total')) || 0, fecha: a('Fecha'), formaPago: a('FormaPago'),
    receptorNombre: a('Nombre', rec), receptorRfc: a('Rfc', rec),
    receptorCP: a('DomicilioFiscalReceptor', rec), receptorUsoCfdi: a('UsoCFDI', rec),
    receptorRegimen: a('RegimenFiscalReceptor', rec),
    uuid: uuidM ? uuidM[1] : ''
  };
}

function _facBuildXmlIndex(fechaInicio, fechaFin) {
  var months = _facMonthsInRange(fechaInicio, fechaFin);
  var byFolio = {}, all = [];
  months.forEach(function (ym) {
    var folder = _facMonthFolder(ym.anio, ym.mes);
    if (!folder) return;
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (!/\.xml$/i.test(name)) continue;
      var xml;
      try { xml = file.getBlob().getDataAsString('UTF-8'); } catch (e) { continue; }
      var p = _facQuickParse(xml);
      if (p.tipo !== 'I') continue; // solo facturas de ingreso (excluye nómina, notas de crédito, pago)
      var fechaISO = p.fecha ? p.fecha.substring(0, 10) : '';
      if (fechaISO < fechaInicio || fechaISO > fechaFin) continue;
      var rec = {
        folio: p.folio, serie: p.serie, uuid: p.uuid, total: p.total, fecha: fechaISO,
        receptorNombre: p.receptorNombre, receptorRfc: p.receptorRfc,
        receptorCP: p.receptorCP, receptorUsoCfdi: p.receptorUsoCfdi, receptorRegimen: p.receptorRegimen,
        formaPago: p.formaPago,
        fileId: file.getId(), fileUrl: file.getUrl(), fileName: name
      };
      all.push(rec);
      if (p.folio) byFolio[String(p.folio)] = rec;
    }
  });
  return { byFolio: byFolio, all: all };
}

function _facReadOpsInRange(fechaInicio, fechaFin) {
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
        id: op, fecha: fecha, paciente: String(r[3] || ''), total: 0,
        factura: String(r[17] || '').trim(), poliza: String(r[18] || ''), archivoURL: String(r[22] || '')
      };
      order.push(op);
    }
    opsMap[op].total += num(r[9]);
  }
  return order.map(function (k) { return opsMap[k]; });
}

function reconciliarFacturasXml(fechaInicio, fechaFin) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    var ops = _facReadOpsInRange(fechaInicio, fechaFin);
    var idx = _facBuildXmlIndex(fechaInicio, fechaFin);
    var conDocumento = [], faltaDocumento = [], sinFactura = [];
    var usedFolios = {};
    ops.forEach(function (op) {
      if (op.archivoURL) { conDocumento.push(op); if (op.factura) usedFolios[op.factura] = true; return; }
      if (op.factura && idx.byFolio[op.factura]) {
        var x = idx.byFolio[op.factura];
        usedFolios[op.factura] = true;
        faltaDocumento.push({
          id: op.id, fecha: op.fecha, paciente: op.paciente, total: op.total, factura: op.factura,
          xmlFileId: x.fileId, xmlFileUrl: x.fileUrl, xmlTotal: x.total, xmlUuid: x.uuid
        });
        return;
      }
      sinFactura.push(op);
    });
    var xmlHuerfanos = idx.all.filter(function (x) { return x.folio && !usedFolios[x.folio]; });
    return {
      ok: true, conDocumento: conDocumento, faltaDocumento: faltaDocumento,
      sinFactura: sinFactura, xmlHuerfanos: xmlHuerfanos, totalOps: ops.length
    };
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
    var data = sheet.getDataRange().getValues();
    var urlCol = BD_INGRESOS_HEADERS.indexOf('ArchivoURL') + 1;
    var updated = 0;
    for (var ri = 1; ri < data.length; ri++) {
      if (String(data[ri][0]) === opId) { sheet.getRange(ri + 1, urlCol).setValue(url); updated++; }
    }
    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (e) {}
    logAudit(body.usuario || 'sistema', 'Ingreso', 'VincularXML', opId, 'ArchivoURL', '', url);
    return { ok: true, url: url, updated: updated };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Parseo completo (XmlService) para el reporte — incluye impuestos por concepto
function _facParseCfdiFull(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var content = file.getBlob().getDataAsString('UTF-8');
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
          ivaPct: ivaPct, ivaMonto: ivaMonto
        });
      }
    }
    return {
      ok: true, serie: attr(root, 'Serie'), folio: attr(root, 'Folio'), fecha: attr(root, 'Fecha'),
      subTotal: parseFloat(attr(root, 'SubTotal')) || 0, total: parseFloat(attr(root, 'Total')) || 0,
      moneda: attr(root, 'Moneda') || 'MXN', tipoCambio: parseFloat(attr(root, 'TipoCambio')) || 1,
      formaPago: attr(root, 'FormaPago'), metodoPago: attr(root, 'MetodoPago'),
      totalImpuestosTrasladados: (function () { var im = child(root, 'Impuestos'); return im ? (parseFloat(attr(im, 'TotalImpuestosTrasladados')) || 0) : 0; })(),
      emisor: { rfc: emisorEl ? attr(emisorEl, 'Rfc') : '', nombre: emisorEl ? attr(emisorEl, 'Nombre') : '' },
      receptor: { rfc: receptorEl ? attr(receptorEl, 'Rfc') : '', nombre: receptorEl ? attr(receptorEl, 'Nombre') : '' },
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
    var facturables = ops.filter(function (o) { return o.factura && idx.byFolio[o.factura]; });
    if (!facturables.length) return { ok: false, error: 'No hay operaciones con factura y XML vinculado en el periodo seleccionado' };

    var pacFiscal = _pacFiscalIndex();

    var maxConceptos = 1;
    var invoices = [];
    facturables.forEach(function (op) {
      var x = idx.byFolio[op.factura];
      var full = _facParseCfdiFull(x.fileId);
      if (!full || !full.ok || !full.conceptos.length) return;
      if (full.conceptos.length > maxConceptos) maxConceptos = full.conceptos.length;
      var pf = pacFiscal[_pacNormNombre(op.paciente)];
      full._paciente = op.paciente;
      full._opId = op.id;
      // Regla: datos fiscales del Registro de Pacientes tienen prioridad; si el
      // paciente no tiene RFC/Razón Social capturados, se usa público en general.
      full._rfcFinal = (pf && pf.rfc) ? pf.rfc : 'XAXX010101000';
      full._razonSocialFinal = (pf && pf.razonSocial) ? pf.razonSocial : 'PUBLICO EN GENERAL';
      invoices.push(full);
    });
    if (!invoices.length) return { ok: false, error: 'No se pudo leer ningún XML vinculado en el periodo' };

    var headers = FAC_MASIVA_HEADERS_INV.slice();
    for (var c = 0; c < maxConceptos; c++) headers = headers.concat(FAC_MASIVA_HEADERS_CONCEPTO);

    var rows = [headers];
    invoices.forEach(function (f) {
      var metodoLabel = FAC_METODO_PAGO_MAP[f.formaPago] || f.formaPago || '';
      var row = [
        f.serie || '', f.fecha ? f.fecha.substring(0, 10) : '', '', f._rfcFinal,
        f.moneda || 'MXN', f.tipoCambio || 1, metodoLabel, '', '',
        f.total, 0, f.subTotal, f.totalImpuestosTrasladados || 0, 0, 0, 0, f.total
      ];
      f.conceptos.forEach(function (con) {
        row.push(
          con.noIdentificacion || con.claveProdServ || '', '', con.cantidad, con.valorUnitario, 0,
          con.ivaPct, con.ivaMonto, 0, 0, 0, 0, 0, 0, 0
        );
      });
      var missing = maxConceptos - f.conceptos.length;
      for (var mi = 0; mi < missing; mi++) row = row.concat(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      rows.push(row);
    });

    // Se escribe el .xlsx directamente en formato OOXML (sin pasar por Sheets ni
    // por conversiones de Drive) para no depender de getAs()/UrlFetchApp, que han
    // fallado en este proyecto. Solo usa Utilities, que ya está autorizado.
    var xlsxBlob = _buildXlsxBlob(rows, 'Hoja1', 'ContaDigital_Masiva_' + fechaInicio + '_a_' + fechaFin + '.xlsx');
    var folder = DriveApp.getFolderById(INGRESOS_FOLDER_FACTURAS);
    var xlsxFile = folder.createFile(xlsxBlob);

    var detalle = invoices.map(function (f) {
      return {
        opId: f._opId, paciente: f._paciente, razonSocial: f._razonSocialFinal,
        rfc: f._rfcFinal, folio: f.folio, serie: f.serie,
        fecha: f.fecha ? f.fecha.substring(0, 10) : '', total: f.total
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
      var x = idx.byFolio[op.factura];
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
      totalXmlEnRango: idx.all.length, totalOpsConFactura: opsConFactura,
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
