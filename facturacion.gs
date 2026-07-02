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
    total: parseFloat(a('Total')) || 0, fecha: a('Fecha'),
    receptorNombre: a('Nombre', rec), receptorRfc: a('Rfc', rec),
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

function generarReporteContaDigital(fechaInicio, fechaFin, usuario) {
  try {
    if (!fechaInicio || !fechaFin) return { ok: false, error: 'Rango de fechas requerido' };
    var ops = _facReadOpsInRange(fechaInicio, fechaFin);
    var idx = _facBuildXmlIndex(fechaInicio, fechaFin);
    var facturables = ops.filter(function (o) { return o.factura && idx.byFolio[o.factura]; });
    if (!facturables.length) return { ok: false, error: 'No hay operaciones con factura y XML vinculado en el periodo seleccionado' };

    var maxConceptos = 1;
    var invoices = [];
    facturables.forEach(function (op) {
      var x = idx.byFolio[op.factura];
      var full = _facParseCfdiFull(x.fileId);
      if (!full || !full.ok || !full.conceptos.length) return;
      if (full.conceptos.length > maxConceptos) maxConceptos = full.conceptos.length;
      invoices.push(full);
    });
    if (!invoices.length) return { ok: false, error: 'No se pudo leer ningún XML vinculado en el periodo' };

    var headers = FAC_MASIVA_HEADERS_INV.slice();
    for (var c = 0; c < maxConceptos; c++) headers = headers.concat(FAC_MASIVA_HEADERS_CONCEPTO);

    var rows = [headers];
    invoices.forEach(function (f) {
      var metodoLabel = FAC_METODO_PAGO_MAP[f.formaPago] || f.formaPago || '';
      var row = [
        f.serie || '', f.fecha ? f.fecha.substring(0, 10) : '', '', f.receptor.rfc || 'XAXX010101000',
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

    var ssOut = SpreadsheetApp.create('ContaDigital_Masiva_' + fechaInicio + '_a_' + fechaFin);
    var sh = ssOut.getSheets()[0];
    sh.setName('Hoja1');
    sh.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    // DriveApp.getAs(MICROSOFT_EXCEL) falla intermitentemente para Sheets nativos;
    // se usa la URL de exportación de Sheets (más confiable) con el token OAuth del script.
    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ssOut.getId() + '/export?format=xlsx';
    var exportResp = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    var xlsxBlob = exportResp.getBlob().setName('ContaDigital_Masiva_' + fechaInicio + '_a_' + fechaFin + '.xlsx');
    var folder = DriveApp.getFolderById(INGRESOS_FOLDER_FACTURAS);
    var xlsxFile = folder.createFile(xlsxBlob);
    DriveApp.getFileById(ssOut.getId()).setTrashed(true);

    logAudit(usuario || 'sistema', 'Facturacion', 'ReporteContaDigital', '', fechaInicio + ' a ' + fechaFin, '', invoices.length + ' facturas');
    return {
      ok: true, url: xlsxFile.getUrl(), numFacturas: invoices.length,
      numOperacionesTotal: ops.length, numSinXml: ops.length - facturables.length
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
