/* ══════════════════════════════════════════════════════════════
   comprobantes.gs — Centro de Comprobantes (Egresos)
   ------------------------------------------------------------
   Indexa la carpeta de Drive donde se guardan por año/mes todos
   los comprobantes (facturas XML de proveedores y sus PDF) y los
   cruza contra Egresos para vincularlos, igual que el sistema de
   facturas de Ingresos:
     Raíz → 2026 → junio / 07 JUL / etc. → archivos .xml/.pdf
   El nombre de la carpeta de mes es flexible: "junio", "06",
   "06 JUN", "6 junio"… todos se reconocen.
   Estados de cada XML:
     vinculado       → un egreso ya apunta a este archivo (Link Factura)
     sugerido        → hay exactamente UN egreso candidato (RFC/proveedor
                       + monto + fecha cercana) — un clic para confirmar
     multiple        → varios candidatos; el usuario elige en el dropdown
     sinCoincidencia → ningún egreso coincide todavía
   Requiere: finance.gs (readEgresosData, EGRESOS_IDS, logAudit),
   providers.gs (readProveedores) y core.gs/finance.gs para el wiring.
   ══════════════════════════════════════════════════════════════ */

/* Dos fuentes, organizadas Año → Mes:
     facturas → repositorio de XML CFDI de egresos del SAT (carpeta 1rIW…,
                donde se descargan por año/mes TODOS los XML de proveedores;
                distinta de "Facturas Recibidas" 1t8--…, que es el destino de
                los adjuntos manuales 📄 y no se toca)
     pagos    → Contabilidad\Pagos (comprobantes de transferencia,
                = EGRESOS_DRIVE_PAGOS de finance.gs) */
var COMP_FACTURAS_ROOT_ID = '1rIWggcMKPAtCRvRxBrQgYzSaCK6kp63w';
var COMP_PAGOS_ROOT_ID    = '1D9H3nNIrkgg2wqJtKXzhuSLDH6hIUoPk';
function _compRootId(fuente) { return fuente === 'pagos' ? COMP_PAGOS_ROOT_ID : COMP_FACTURAS_ROOT_ID; }
var COMP_MESES_NOMBRES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
var COMP_MESES_ABR = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];

function _compMesDesdeNombre(nombre) {
  var n = String(nombre || '').toLowerCase().trim();
  var m = n.match(/^(\d{1,2})\b/);
  if (m) { var num = parseInt(m[1], 10); if (num >= 1 && num <= 12) return num; }
  for (var i = 0; i < 12; i++) {
    if (n.indexOf(COMP_MESES_NOMBRES[i]) > -1) return i + 1;
  }
  for (var j = 0; j < 12; j++) {
    if (n.indexOf(COMP_MESES_ABR[j]) > -1) return j + 1;
  }
  return 0;
}

/* ── Estructura: qué años y meses existen en la carpeta ─────────── */
function listComprobantesEstructura(fuente) {
  try {
    var root = DriveApp.getFolderById(_compRootId(fuente));
    var anios = [];
    var it = root.getFolders();
    while (it.hasNext()) {
      var f = it.next();
      var y = f.getName().match(/20\d{2}/);
      if (!y) continue;
      var meses = [];
      var mit = f.getFolders();
      while (mit.hasNext()) {
        var mf = mit.next();
        var mesNum = _compMesDesdeNombre(mf.getName());
        if (mesNum > 0) meses.push({ carpeta: mf.getName(), mes: mesNum, folderId: mf.getId() });
      }
      meses.sort(function (a, b) { return a.mes - b.mes; });
      anios.push({ anio: parseInt(y[0], 10), folderId: f.getId(), meses: meses });
    }
    anios.sort(function (a, b) { return b.anio - a.anio; });
    return { ok: true, anios: anios, rootUrl: root.getUrl() };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Parser ligero de CFDI por regex (sin XmlService, para poder
   indexar un mes completo en una sola llamada sin agotar tiempo).
   El detalle completo (conceptos, sellos…) lo da leerXmlFactura(). ── */
function _compParseXmlLight(content) {
  function attrDe(tag, name) {
    var re = new RegExp('<[A-Za-z0-9]*:?' + tag + '\\b[^>]*?\\s' + name + '="([^"]*)"');
    var m = content.match(re);
    return m ? m[1] : '';
  }
  return {
    serie: attrDe('Comprobante', 'Serie'),
    folio: attrDe('Comprobante', 'Folio'),
    fecha: attrDe('Comprobante', 'Fecha'),
    total: parseFloat(attrDe('Comprobante', 'Total')) || 0,
    subTotal: parseFloat(attrDe('Comprobante', 'SubTotal')) || 0,
    moneda: attrDe('Comprobante', 'Moneda') || 'MXN',
    tipoComprobante: attrDe('Comprobante', 'TipoDeComprobante') || 'I',
    metodoPago: attrDe('Comprobante', 'MetodoPago'),
    emisorRfc: attrDe('Emisor', 'Rfc'),
    emisorNombre: attrDe('Emisor', 'Nombre'),
    receptorRfc: attrDe('Receptor', 'Rfc'),
    uuid: attrDe('TimbreFiscalDigital', 'UUID')
  };
}

/* ── Leer un mes de una fuente:
   'facturas' → XMLs CFDI parseados + PDF hermano + cruce Link Factura
   'pagos'    → cualquier archivo (PDF/imagen) + cruce Link Pago ─────── */
function readComprobantesMes(anio, mes, fuente) {
  try {
    fuente = fuente === 'pagos' ? 'pagos' : 'facturas';
    anio = parseInt(anio, 10) || new Date().getFullYear();
    mes = parseInt(mes, 10) || (new Date().getMonth() + 1);

    var est = listComprobantesEstructura(fuente);
    if (!est.ok) return est;
    var anioObj = null;
    for (var a = 0; a < est.anios.length; a++) if (est.anios[a].anio === anio) { anioObj = est.anios[a]; break; }
    if (!anioObj) return { ok: true, fuente: fuente, anio: anio, mes: mes, comprobantes: [], estructura: est.anios, aviso: 'No existe carpeta para el año ' + anio };
    var mesObj = null;
    for (var m = 0; m < anioObj.meses.length; m++) if (anioObj.meses[m].mes === mes) { mesObj = anioObj.meses[m]; break; }
    if (!mesObj) return { ok: true, fuente: fuente, anio: anio, mes: mes, comprobantes: [], estructura: est.anios, aviso: 'No existe carpeta del mes ' + mes + ' en ' + anio };

    if (fuente === 'pagos') return _compLeerMesPagos(anio, mes, mesObj, est);

    // 1. Archivos del mes: XMLs parseados; PDFs guardados por nombre base
    var folder = DriveApp.getFolderById(mesObj.folderId);
    var files = folder.getFiles();
    var xmls = [];
    var pdfsPorBase = {};
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (/\.xml$/i.test(name)) {
        var parsed;
        try { parsed = _compParseXmlLight(file.getBlob().getDataAsString('UTF-8')); }
        catch (pe) { parsed = { error: pe.message }; }
        xmls.push({
          fileId: file.getId(), fileName: name, url: file.getUrl(),
          serie: parsed.serie || '', folio: parsed.folio || '', fecha: (parsed.fecha || '').substring(0, 10),
          total: parsed.total || 0, moneda: parsed.moneda || 'MXN', tipoComprobante: parsed.tipoComprobante || '',
          metodoPago: parsed.metodoPago || '', emisorRfc: (parsed.emisorRfc || '').toUpperCase(),
          emisorNombre: parsed.emisorNombre || '', receptorRfc: parsed.receptorRfc || '',
          uuid: (parsed.uuid || '').toUpperCase(), parseError: parsed.error || ''
        });
      } else if (/\.pdf$/i.test(name)) {
        pdfsPorBase[name.replace(/\.pdf$/i, '').toLowerCase().trim()] = { fileId: file.getId(), url: file.getUrl(), fileName: name };
      }
    }
    // PDF hermano: mismo nombre base que el XML
    xmls.forEach(function (x) {
      var base = x.fileName.replace(/\.xml$/i, '').toLowerCase().trim();
      if (pdfsPorBase[base]) { x.pdfFileId = pdfsPorBase[base].fileId; x.pdfUrl = pdfsPorBase[base].url; }
    });

    // 2. Egresos del año + mapa RFC→proveedor del catálogo
    var eg = readEgresosData(anio);
    var egRows = (eg.ok && eg.rows) ? eg.rows : [];
    var provPorRfc = {};
    try {
      var provs = readProveedores();
      (provs.rows || []).forEach(function (p) {
        var rfc = String(p.rfc || '').toUpperCase().replace(/[\s-]/g, '');
        if (rfc) provPorRfc[rfc] = String(p.nombre || '');
      });
    } catch (pe2) {}

    function norm(s) { return String(s || '').toLowerCase().replace(/[\s.,]+/g, ' ').trim(); }
    function diasEntre(f1, f2) {
      try { return Math.abs((new Date(f1 + 'T12:00:00') - new Date(f2 + 'T12:00:00')) / 86400000); }
      catch (e) { return 9999; }
    }

    // 3. Clasificación de cada XML contra los egresos
    xmls.forEach(function (x) {
      x.estado = 'sinCoincidencia';
      x.candidatos = [];

      // ¿Algún egreso ya apunta a este archivo?
      for (var i = 0; i < egRows.length; i++) {
        var lk = (egRows[i].linkFacturaUrl || '') + ' ' + (egRows[i].linkFactura || '');
        if (x.fileId && lk.indexOf(x.fileId) > -1) {
          x.estado = 'vinculado';
          x.egreso = { rowNum: egRows[i]._rowNum, proveedor: egRows[i].proveedor, concepto: egRows[i].concepto, monto: egRows[i].monto, fecha: egRows[i].fecha, pagado: egRows[i].pagado };
          break;
        }
      }
      if (x.estado === 'vinculado') return;

      // Candidatos: monto igual (±$0.50), sin factura vinculada aún, fecha cercana (±45 días).
      // Prioridad a los del proveedor cuyo RFC coincide con el emisor del XML.
      var provDelRfc = provPorRfc[x.emisorRfc.replace(/[\s-]/g, '')] || '';
      var nombreEmisor = norm(x.emisorNombre);
      for (var j = 0; j < egRows.length; j++) {
        var r = egRows[j];
        if (r.linkFacturaUrl || r.linkFactura) continue;
        if (Math.abs((r.monto || 0) - x.total) > 0.5 || !x.total) continue;
        var fechaRef = r.fecha || r.vencimiento || '';
        if (x.fecha && fechaRef && diasEntre(x.fecha, fechaRef) > 45) continue;
        var provNorm = norm(r.proveedor);
        var matchProv = (provDelRfc && norm(provDelRfc) === provNorm)
          || (nombreEmisor && provNorm && (nombreEmisor.indexOf(provNorm) > -1 || provNorm.indexOf(nombreEmisor) > -1));
        x.candidatos.push({
          rowNum: r._rowNum, proveedor: r.proveedor, concepto: r.concepto,
          monto: r.monto, fecha: fechaRef, pagado: r.pagado, matchProveedor: !!matchProv
        });
      }
      // Si hay candidatos con match de proveedor, esos mandan
      var fuertes = x.candidatos.filter(function (c) { return c.matchProveedor; });
      if (fuertes.length === 1) { x.estado = 'sugerido'; x.candidatos = fuertes; }
      else if (fuertes.length > 1) { x.estado = 'multiple'; x.candidatos = fuertes; }
      else if (x.candidatos.length === 1) { x.estado = 'sugerido'; }
      else if (x.candidatos.length > 1) { x.estado = 'multiple'; }
      x.candidatos = x.candidatos.slice(0, 10);
    });

    xmls.sort(function (a, b) { return (a.fecha < b.fecha) ? 1 : -1; });

    // Egresos del mes que siguen sin factura (para el resumen)
    var mesStr = String(anio) + '-' + String(mes).padStart(2, '0');
    var egresosSinFactura = egRows.filter(function (r) {
      var f = r.fecha || r.vencimiento || '';
      return f.indexOf(mesStr) === 0 && !r.linkFacturaUrl && !r.linkFactura && (r.monto || 0) > 0;
    }).length;

    return {
      ok: true, fuente: 'facturas', anio: anio, mes: mes, carpeta: mesObj.carpeta,
      comprobantes: xmls, estructura: est.anios,
      resumen: {
        total: xmls.length,
        vinculados: xmls.filter(function (x) { return x.estado === 'vinculado'; }).length,
        sugeridos: xmls.filter(function (x) { return x.estado === 'sugerido' || x.estado === 'multiple'; }).length,
        sinCoincidencia: xmls.filter(function (x) { return x.estado === 'sinCoincidencia'; }).length,
        egresosSinFactura: egresosSinFactura
      }
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Fuente "pagos": comprobantes de transferencia/pago (PDF, imagen…).
   No hay XML que parsear, así que el cruce es directo por Link Pago y
   el vinculado manual ofrece los egresos PAGADOS que aún no tienen
   comprobante (cualquier mes — un pago puede cubrir meses previos). ── */
function _compLeerMesPagos(anio, mes, mesObj, est) {
  var folder = DriveApp.getFolderById(mesObj.folderId);
  var files = folder.getFiles();
  var archivos = [];
  while (files.hasNext()) {
    var f = files.next();
    archivos.push({
      fileId: f.getId(), fileName: f.getName(), url: f.getUrl(),
      fecha: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
    });
  }
  var eg = readEgresosData(anio);
  var egRows = (eg.ok && eg.rows) ? eg.rows : [];
  archivos.forEach(function (x) {
    x.estado = 'sinVincular';
    for (var i = 0; i < egRows.length; i++) {
      var lk = (egRows[i].linkPagoUrl || '') + ' ' + (egRows[i].linkPago || '');
      if (lk.indexOf(x.fileId) > -1) {
        x.estado = 'vinculado';
        x.egreso = { rowNum: egRows[i]._rowNum, proveedor: egRows[i].proveedor, concepto: egRows[i].concepto, monto: egRows[i].monto, fecha: egRows[i].fecha };
        break;
      }
    }
  });
  archivos.sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });

  var mesStr = String(anio) + '-' + String(mes).padStart(2, '0');
  var candidatos = egRows.filter(function (r) {
    return r.pagado && !r.linkPagoUrl && !r.linkPago && (r.monto || 0) > 0;
  }).map(function (r) {
    return { rowNum: r._rowNum, proveedor: r.proveedor, concepto: r.concepto, monto: r.monto, fecha: r.fecha };
  });
  candidatos.sort(function (a, b) { return (a.fecha || '') < (b.fecha || '') ? 1 : -1; });

  return {
    ok: true, fuente: 'pagos', anio: anio, mes: mes, carpeta: mesObj.carpeta,
    comprobantes: archivos, estructura: est.anios,
    egresosCandidatos: candidatos.slice(0, 300),
    resumen: {
      total: archivos.length,
      vinculados: archivos.filter(function (x) { return x.estado === 'vinculado'; }).length,
      sinVincular: archivos.filter(function (x) { return x.estado !== 'vinculado'; }).length,
      egresosPagadosSinComprobante: candidatos.filter(function (c) { return (c.fecha || '').indexOf(mesStr) === 0; }).length
    }
  };
}

/* ── Vincular: escribe el hipervínculo del archivo en la columna del
   egreso elegido — campo 'factura' → Link Factura, 'pago' → Link Pago ── */
function vincularComprobanteEgreso(body) {
  try {
    var anio = parseInt(body.anio, 10) || new Date().getFullYear();
    var rowNum = parseInt(body.rowNum, 10);
    var fileId = String(body.fileId || '').trim();
    var campo = body.campo === 'pago' ? 'pago' : 'factura';
    if (!rowNum || !fileId) return { ok: false, error: 'Faltan rowNum o fileId' };

    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = (typeof EGRESOS_TABS !== 'undefined' && EGRESOS_TABS[anio]) ? EGRESOS_TABS[anio] : 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(tabName) || ss.getSheets()[0];

    var buscar = campo === 'pago' ? 'link pago' : 'link factura';
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var iCol = -1;
    for (var c = 0; c < headers.length; c++) if (headers[c].indexOf(buscar) > -1) { iCol = c; break; }
    if (iCol < 0) return { ok: false, error: 'No se encontró la columna "' + buscar + '" en ' + tabName };

    var url = 'https://drive.google.com/file/d/' + fileId + '/view';
    var etiqueta = body.etiqueta || (campo === 'pago' ? 'Comprobante de pago' : 'Factura XML');
    var rich = SpreadsheetApp.newRichTextValue().setText(etiqueta).setLinkUrl(url).build();
    sheet.getRange(rowNum, iCol + 1).setRichTextValue(rich);

    try { logAudit(body.usuario || 'sistema', 'Comprobantes', 'Vincular' + (campo === 'pago' ? 'Pago' : 'Factura'), 'fila ' + rowNum, '', '', (body.uuid || fileId)); } catch (e) {}
    return { ok: true, rowNum: rowNum, url: url, campo: campo };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Vincular sugeridos en lote: toma todos los XML del mes cuyo estado
   sea "sugerido" (UN candidato claro; los egresos con factura adjunta ya
   quedaron fuera del cruce) y escribe cada hipervínculo de golpe —
   igual que el "Vincular todo" de la Facturación de Ingresos. ──────── */
function vincularComprobantesLote(body) {
  try {
    var anio = parseInt(body.anio, 10) || new Date().getFullYear();
    var mes = parseInt(body.mes, 10) || (new Date().getMonth() + 1);
    var data = readComprobantesMes(anio, mes, 'facturas');
    if (!data.ok) return data;
    var vinculados = 0, detalles = [], errores = [];
    (data.comprobantes || []).forEach(function (x) {
      if (x.estado !== 'sugerido' || !x.candidatos || !x.candidatos.length) return;
      var c = x.candidatos[0];
      var r = vincularComprobanteEgreso({
        anio: anio, rowNum: c.rowNum, fileId: x.fileId, campo: 'factura',
        uuid: x.uuid, etiqueta: (x.serie ? x.serie + '-' : '') + (x.folio || 'Factura XML'),
        usuario: body.usuario || 'sistema'
      });
      if (r.ok) { vinculados++; detalles.push({ folio: (x.serie ? x.serie + '-' : '') + (x.folio || ''), proveedor: c.proveedor, monto: c.monto }); }
      else errores.push((x.folio || x.fileName) + ': ' + r.error);
    });
    try { logAudit(body.usuario || 'sistema', 'Comprobantes', 'VincularLote', anio + '-' + mes, '', '', vinculados + ' vinculados'); } catch (e) {}
    return { ok: true, vinculados: vinculados, detalles: detalles, errores: errores };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Menú (estructura aprobada por Enrique): "Comprobantes" DENTRO de
   Egresos — primero Gastos Operativos (fin-gastos, la vista de Egresos)
   y Comprobantes colgando de él. El sidebar renderiza la vista-con-hijos
   como navegable + desplegable. Idempotente: si la fila ya está bien,
   no toca nada; si tiene otro padre, la regresa a fin-gastos. ───────── */
function configurarMenuComprobantes() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Menu');
  if (!sh) return { ok: false, error: 'No se encontró la hoja Menu' };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var idxCol = hdrs.indexOf('ID'), padreCol = hdrs.indexOf('Padre');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxCol] || '').trim() === 'comprobantes') {
      if (String(data[i][padreCol] || '').trim() !== 'fin-gastos') {
        sh.getRange(i + 1, padreCol + 1).setValue('fin-gastos');
        return { ok: true, aviso: 'La entrada "comprobantes" se movió dentro de Egresos (fin-gastos)' };
      }
      return { ok: true, aviso: 'Ya está bien: comprobantes dentro de Egresos (fin-gastos)' };
    }
  }
  var fila = ['comprobantes', 'fin-gastos', 'Comprobantes', '', 'paperclip', 1, 'vista', 'comprobantes', 'TRUE'];
  sh.getRange(sh.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);
  return { ok: true, creada: true };
}

/* ══════════════════════════════════════════════════════════════
   ENVÍO DE DOCUMENTOS POR CORREO
   ------------------------------------------------------------
   Plantillas por tipo de documento (hoja Plantillas_Correo en el
   spreadsheet principal, editables también desde el ERP) y envío
   DIRECTO desde la cuenta Google del deployment vía GmailApp, con
   los archivos de Drive adjuntos de verdad. Cada envío queda en
   la bitácora Correos_Enviados y en la carpeta Enviados de Gmail.
   Variables disponibles en asunto/cuerpo:
     {{destinatario}} {{folio}} {{monto}} {{fecha}} {{referencia}}
     {{proveedor}} {{paciente}} {{usuario}}
   ══════════════════════════════════════════════════════════════ */
var CORREO_PLANTILLAS_TAB = 'Plantillas_Correo';
var CORREO_LOG_TAB = 'Correos_Enviados';

var CORREO_PLANTILLAS_DEFAULT = [
  ['factura-ingreso', 'Factura de ingreso (paciente)',
   'Factura {{folio}} — Hestia Fertility',
   'Estimado(a) {{destinatario}}:\n\nLe compartimos su factura {{folio}} por {{monto}}, emitida el {{fecha}}.\n\nQuedamos atentos a cualquier aclaración.\n\nHestia Fertility · Administración'],
  ['factura-egreso', 'Factura de egreso (proveedor)',
   'Factura {{folio}} recibida — Hestia Fertility',
   'Estimado proveedor:\n\nConfirmamos la recepción de la factura {{folio}} por {{monto}} con fecha {{fecha}}.\n\nQuedamos atentos a cualquier aclaración.\n\nHestia Fertility · Administración'],
  ['comprobante-pago', 'Comprobante de pago',
   'Comprobante de pago — Hestia Fertility · {{referencia}}',
   'Estimado proveedor:\n\nPor este medio le compartimos el comprobante de pago correspondiente a {{referencia}} por {{monto}}, liquidado el {{fecha}}.\n\nQuedamos atentos a cualquier aclaración.\n\nHestia Fertility · Administración'],
  ['cotizacion', 'Cotización',
   'Cotización — Hestia Fertility · {{referencia}}',
   'Estimado(a) {{destinatario}}:\n\nAdjuntamos la cotización solicitada ({{referencia}}).\n\nQuedamos a sus órdenes para cualquier duda.\n\nHestia Fertility · Administración'],
  ['estado-cuenta', 'Estado de cuenta',
   'Estado de cuenta — Hestia Fertility',
   'Estimado(a) {{destinatario}}:\n\nLe compartimos su estado de cuenta al {{fecha}}.\n\nQuedamos atentos a cualquier aclaración.\n\nHestia Fertility · Administración']
];

function setupPlantillasCorreo() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CORREO_PLANTILLAS_TAB);
  if (!sh) {
    sh = ss.insertSheet(CORREO_PLANTILLAS_TAB);
    sh.getRange(1, 1, 1, 4).setValues([['ID', 'Nombre', 'Asunto', 'Cuerpo']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f4f6');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, CORREO_PLANTILLAS_DEFAULT.length, 4).setValues(CORREO_PLANTILLAS_DEFAULT);
  }
  var log = ss.getSheetByName(CORREO_LOG_TAB);
  if (!log) {
    log = ss.insertSheet(CORREO_LOG_TAB);
    log.getRange(1, 1, 1, 8).setValues([['Timestamp', 'Usuario', 'Para', 'CC', 'Asunto', 'TipoDocumento', 'Referencia', 'Adjuntos']]);
    log.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
    log.setFrozenRows(1);
  }
  return { ok: true };
}

function readPlantillasCorreo() {
  try {
    setupPlantillasCorreo();
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(CORREO_PLANTILLAS_TAB);
    var data = sh.getDataRange().getValues();
    var plantillas = [];
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][0] || '').trim();
      if (!id) continue;
      plantillas.push({ id: id, nombre: String(data[i][1] || id), asunto: String(data[i][2] || ''), cuerpo: String(data[i][3] || '') });
    }
    return { ok: true, plantillas: plantillas };
  } catch (ex) { return { ok: false, error: ex.message, plantillas: [] }; }
}

function savePlantillaCorreo(body) {
  try {
    setupPlantillasCorreo();
    var id = String(body.id || '').trim();
    if (!id) return { ok: false, error: 'Falta el ID de la plantilla' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(CORREO_PLANTILLAS_TAB);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === id) {
        if (body.nombre !== undefined) sh.getRange(i + 1, 2).setValue(body.nombre);
        if (body.asunto !== undefined) sh.getRange(i + 1, 3).setValue(body.asunto);
        if (body.cuerpo !== undefined) sh.getRange(i + 1, 4).setValue(body.cuerpo);
        try { logAudit(body.usuario || 'sistema', 'Correo', 'EditarPlantilla', id, '', '', ''); } catch (e) {}
        return { ok: true, id: id };
      }
    }
    sh.appendRow([id, body.nombre || id, body.asunto || '', body.cuerpo || '']);
    return { ok: true, id: id, creada: true };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function enviarDocumentoCorreo(body) {
  try {
    var para = String(body.para || '').trim();
    var asunto = String(body.asunto || '').trim();
    var cuerpo = String(body.cuerpo || '');
    if (!para || !/^[^\s@]+@[^\s@]+\.[^\s@]+(\s*[,;]\s*[^\s@]+@[^\s@]+\.[^\s@]+)*$/.test(para))
      return { ok: false, error: 'Destinatario inválido: "' + para + '"' };
    if (!asunto) return { ok: false, error: 'El asunto es obligatorio' };

    var adjuntos = body.adjuntos || [];
    var blobs = [], nombres = [], fallidos = [];
    for (var i = 0; i < Math.min(adjuntos.length, 8); i++) {
      var fid = String(adjuntos[i] || '').trim();
      if (!fid) continue;
      try {
        var f = DriveApp.getFileById(fid);
        blobs.push(f.getBlob());
        nombres.push(f.getName());
      } catch (fe) { fallidos.push(fid); }
    }
    if (fallidos.length && !blobs.length)
      return { ok: false, error: 'No se pudo leer ningún adjunto de Drive (' + fallidos.length + ' fallidos)' };

    var opts = { name: 'Hestia Fertility' };
    if (blobs.length) opts.attachments = blobs;
    var cc = String(body.cc || '').trim();
    if (cc) opts.cc = cc;
    GmailApp.sendEmail(para, asunto, cuerpo, opts);

    // Bitácora
    try {
      setupPlantillasCorreo();
      var log = SpreadsheetApp.openById(SHEET_ID).getSheetByName(CORREO_LOG_TAB);
      log.appendRow([new Date(), body.usuario || '', para, cc, asunto, body.tipoDoc || '', body.referencia || '', nombres.join(', ')]);
    } catch (le) {}
    try { logAudit(body.usuario || 'sistema', 'Correo', 'Enviar', body.referencia || '', '', '', para + ' | ' + asunto); } catch (ae) {}

    return { ok: true, para: para, adjuntos: nombres.length, fallidos: fallidos.length };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
