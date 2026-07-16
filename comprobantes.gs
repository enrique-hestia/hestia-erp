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
    emisorRegimen: attrDe('Emisor', 'RegimenFiscal'),
    lugarExpedicion: attrDe('Comprobante', 'LugarExpedicion'),
    formaPago: attrDe('Comprobante', 'FormaPago'),
    receptorRfc: attrDe('Receptor', 'Rfc'),
    receptorUsoCfdi: attrDe('Receptor', 'UsoCFDI'),
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

    // Egresos del mes que siguen sin factura — lista completa para poder
    // mostrarlos (y resolverlos) desde el propio Centro de Comprobantes
    var mesStr = String(anio) + '-' + String(mes).padStart(2, '0');
    var egresosSinFacturaLista = egRows.filter(function (r) {
      var f = r.fecha || r.vencimiento || '';
      return f.indexOf(mesStr) === 0 && !r.linkFacturaUrl && !r.linkFactura && (r.monto || 0) > 0;
    }).map(function (r) {
      return { rowNum: r._rowNum, fecha: (r.fecha || r.vencimiento || '').substring(0, 10),
               proveedor: r.proveedor, concepto: r.concepto, monto: r.monto, pagado: r.pagado };
    });
    egresosSinFacturaLista.sort(function (a, b) { return (a.fecha || '') < (b.fecha || '') ? 1 : -1; });

    return {
      ok: true, fuente: 'facturas', anio: anio, mes: mes, carpeta: mesObj.carpeta,
      comprobantes: xmls, estructura: est.anios,
      egresosSinFacturaLista: egresosSinFacturaLista.slice(0, 200),
      resumen: {
        total: xmls.length,
        vinculados: xmls.filter(function (x) { return x.estado === 'vinculado'; }).length,
        sugeridos: xmls.filter(function (x) { return x.estado === 'sugerido' || x.estado === 'multiple'; }).length,
        sinCoincidencia: xmls.filter(function (x) { return x.estado === 'sinCoincidencia'; }).length,
        egresosSinFactura: egresosSinFacturaLista.length
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

    // Libro/pestaña del AÑO (cada año es un spreadsheet distinto). _egIdDeAnio /
    // _egTabDeAnio (finance.gs) son el único punto de verdad; no repetir el
    // EGRESOS_IDS[anio]||EGRESOS_SS_2026 a mano, que es como se acaba clavado en 2026.
    var ssId = _egIdDeAnio(anio);
    var tabName = _egTabDeAnio(anio);
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

    // ── Persistir los datos fiscales del CFDI en la fila ──────────────────
    // Antes el UUID solo se mandaba a logAudit y se TIRABA: el folio fiscal
    // quedaba únicamente dentro del XML de Drive, así que no se podía buscar.
    // Se lee del PROPIO archivo (no de body.uuid) porque hay vías que no lo
    // mandan — p.ej. egVincularXmlDetectado, el botón "Vincular" del detalle
    // del egreso, solo manda fileId. body.uuid queda como respaldo por si
    // Drive falla (el lote sí lo trae, del índice del mes).
    var fiscal = null;
    if (campo === 'factura' && typeof _egAplicarFacturaFiscal === 'function') {
      fiscal = _egAplicarFacturaFiscal(sheet, rowNum, fileId, body.uuid);
    }

    try { logAudit(body.usuario || 'sistema', 'Comprobantes', 'Vincular' + (campo === 'pago' ? 'Pago' : 'Factura'), 'fila ' + rowNum, '', '', ((fiscal && fiscal.datos && fiscal.datos.uuid) || body.uuid || fileId)); } catch (e) {}
    var out = { ok: true, rowNum: rowNum, url: url, campo: campo };
    if (fiscal) {
      if (fiscal.datos) out.fiscal = fiscal.datos;
      if (fiscal.avisoDuplicado) out.avisoDuplicado = fiscal.avisoDuplicado;
    }
    return out;
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Índice ligero de un mes de XMLs, cacheado 10 min (CacheService) —
   parsear ~100 XMLs de Drive toma varios segundos; el caché hace que
   las búsquedas por egreso sean casi instantáneas después de la 1a. ── */
function _compIndexMes(anio, mes) {
  var cache = CacheService.getScriptCache();
  var key = 'compidx_' + anio + '_' + mes;
  var cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  var est = listComprobantesEstructura('facturas');
  if (!est.ok) return [];
  var anioObj = null;
  for (var a = 0; a < est.anios.length; a++) if (est.anios[a].anio === anio) { anioObj = est.anios[a]; break; }
  if (!anioObj) return [];
  var mesObj = null;
  for (var m = 0; m < anioObj.meses.length; m++) if (anioObj.meses[m].mes === mes) { mesObj = anioObj.meses[m]; break; }
  if (!mesObj) return [];

  var files = DriveApp.getFolderById(mesObj.folderId).getFiles();
  var idx = [];
  while (files.hasNext()) {
    var f = files.next();
    if (!/\.xml$/i.test(f.getName())) continue;
    var p;
    try { p = _compParseXmlLight(f.getBlob().getDataAsString('UTF-8')); } catch (pe) { continue; }
    idx.push({
      fileId: f.getId(), fileName: f.getName(),
      serie: p.serie || '', folio: p.folio || '', fecha: (p.fecha || '').substring(0, 10),
      total: p.total || 0, emisorNombre: p.emisorNombre || '',
      emisorRfc: (p.emisorRfc || '').toUpperCase(), uuid: (p.uuid || '').toUpperCase()
    });
  }
  try { cache.put(key, JSON.stringify(idx), 600); } catch (ce) {}
  return idx;
}

/* ── Buscar el XML de UN egreso concreto (sin factura adjunta): por
   monto (±$0.50) y proveedor (RFC del catálogo o nombre), en el mes
   del egreso y los vecinos. Lo usa el detalle expandido de Egresos. ── */
function buscarXmlParaEgreso(body) {
  try {
    var monto = Number(body.monto) || 0;
    if (!monto) return { ok: false, error: 'Falta el monto del egreso' };
    var fecha = String(body.fecha || '').substring(0, 10);
    var anio = fecha ? parseInt(fecha.substring(0, 4), 10) : new Date().getFullYear();
    var mesBase = fecha ? parseInt(fecha.substring(5, 7), 10) : (new Date().getMonth() + 1);

    var provRfc = '';
    var provNombre = String(body.proveedor || '').toLowerCase().replace(/[\s.,]+/g, ' ').trim();
    try {
      var provs = readProveedores();
      (provs.rows || []).forEach(function (p) {
        var n = String(p.nombre || '').toLowerCase().replace(/[\s.,]+/g, ' ').trim();
        if (n && provNombre && (n === provNombre || provNombre.indexOf(n) > -1 || n.indexOf(provNombre) > -1)) {
          provRfc = String(p.rfc || '').toUpperCase().replace(/[\s-]/g, '');
        }
      });
    } catch (e) {}

    // REGLA (pedido del usuario): la factura se busca en el MISMO MES del egreso.
    // Solo si NO hay del mismo mes Y el mes YA TERMINÓ se extiende a meses vecinos
    // (por si se facturó/pagó en otro mes). Si el mes sigue EN CURSO no se toma la de
    // otro mes: se deja sin vincular esperando la factura del mes (o el cierre del mes).
    var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    var finMes = new Date(anio, mesBase, 0);           // último día del mes del egreso
    var mesTerminado = finMes < hoy;

    function buscarEnMes(y, m, dist) {
      var out = [];
      var yy = y, mm = m; if (mm < 1) { mm = 12; yy--; } if (mm > 12) { mm = 1; yy++; }
      _compIndexMes(yy, mm).forEach(function (x) {
        if (Math.abs((x.total || 0) - monto) > 0.5) return;
        var rfcOk = provRfc && x.emisorRfc.replace(/[\s-]/g, '') === provRfc;
        var nomEmisor = String(x.emisorNombre || '').toLowerCase().replace(/[\s.,]+/g, ' ').trim();
        var nombreOk = provNombre && nomEmisor && (nomEmisor.indexOf(provNombre) > -1 || provNombre.indexOf(nomEmisor) > -1);
        out.push(Object.assign({}, x, { matchProveedor: !!(rfcOk || nombreOk), mesDist: dist, mismoMes: dist === 0 }));
      });
      return out;
    }

    var candidatos = buscarEnMes(anio, mesBase, 0);     // 1) mismo mes
    if (!candidatos.length) {
      if (!mesTerminado) {
        // 3) el mes NO ha terminado → no obviar; esperar la factura del mes
        return { ok: true, encontrados: [], soloMonto: false, esperandoMes: true,
                 mes: anio + '-' + ('0' + mesBase).slice(-2) };
      }
      // 2) mes cerrado sin factura del mismo mes → extender a meses vecinos
      candidatos = buscarEnMes(anio, mesBase - 1, 1).concat(buscarEnMes(anio, mesBase + 1, 1));
    }
    candidatos.sort(function (a, b) { return ((b.matchProveedor ? 1 : 0) - (a.matchProveedor ? 1 : 0)) || (a.mesDist - b.mesDist); });
    var fuertes = candidatos.filter(function (c) { return c.matchProveedor; });
    return { ok: true, encontrados: (fuertes.length ? fuertes : candidatos).slice(0, 5),
             soloMonto: !fuertes.length, mesTerminado: mesTerminado };
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
    var vinculados = 0, detalles = [], errores = [], avisos = [];
    (data.comprobantes || []).forEach(function (x) {
      if (x.estado !== 'sugerido' || !x.candidatos || !x.candidatos.length) return;
      var c = x.candidatos[0];
      var r = vincularComprobanteEgreso({
        anio: anio, rowNum: c.rowNum, fileId: x.fileId, campo: 'factura',
        uuid: x.uuid, etiqueta: (x.serie ? x.serie + '-' : '') + (x.folio || 'Factura XML'),
        usuario: body.usuario || 'sistema'
      });
      if (r.ok) {
        vinculados++;
        detalles.push({ folio: (x.serie ? x.serie + '-' : '') + (x.folio || ''), proveedor: c.proveedor, monto: c.monto });
        // El aviso de CFDI repetido NO se traga: en un lote es justo donde más
        // fácil pasa desapercibido que dos egresos quedaron con la misma factura.
        if (r.avisoDuplicado) avisos.push((x.folio || x.fileName) + ': ' + r.avisoDuplicado);
      }
      else errores.push((x.folio || x.fileName) + ': ' + r.error);
    });
    try { logAudit(body.usuario || 'sistema', 'Comprobantes', 'VincularLote', anio + '-' + mes, '', '', vinculados + ' vinculados'); } catch (e) {}
    return { ok: true, vinculados: vinculados, detalles: detalles, errores: errores, avisos: avisos };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ══════════════════════════════════════════════════════════════
   BACK-FILL de datos fiscales de Egresos ya vinculados
   ------------------------------------------------------------
   Los egresos que YA tenían su factura vinculada antes del build
   .209 traen el UUID dentro del XML de Drive, pero no en la hoja:
   no se pueden buscar. Esto los recorre, abre el XML y llena
   FacturaUUID/FacturaFolio/FacturaRFC/FacturaRazonSocial.

   REANUDABLE por presupuesto de tiempo (mismo patrón que
   _provEmisorIndex, providers.gs): Apps Script corta a los 6 min y
   abrir ~1 XML de Drive cuesta ~0.2-0.5 s, así que un año con
   cientos de facturas NO cabe en una corrida. En vez de morir a la
   mitad en silencio, se corta solo y devuelve `siguiente` (la fila
   por la que va) y `faltan`. Se vuelve a llamar con `desde:
   siguiente` hasta que `truncado` sea false.

   IDEMPOTENTE: una fila que ya tiene UUID se salta (cuenta en
   `yaListos`), salvo que se pida `sobrescribir:true`.
   ══════════════════════════════════════════════════════════════ */
function backfillFacturasEgresos(body) {
  try {
    body = body || {};
    var t0 = Date.now();
    var anio = parseInt(body.anio, 10) || new Date().getFullYear();
    var maxMs = parseInt(body.maxMs, 10) || 120000;   // 2 min: deja aire de sobra bajo el corte de 6
    var desde = parseInt(body.desde, 10) || 2;        // fila 2 = primer dato (1 = headers)
    var sobrescribir = (body.sobrescribir === true || body.sobrescribir === 'true');
    if (desde < 2) desde = 2;

    if (typeof _egIdDeAnio !== 'function' || typeof _egFacturaCols !== 'function')
      return { ok:false, error:'Falta finance.gs (o está desactualizado) en el proyecto de Apps Script. Redespliega.' };
    if (typeof EGRESOS_IDS !== 'undefined' && !EGRESOS_IDS[anio])
      return { ok:false, error:'Año no configurado en EGRESOS_IDS: ' + anio };

    var ss = SpreadsheetApp.openById(_egIdDeAnio(anio));   // libro DEL AÑO
    var sh = ss.getSheetByName(_egTabDeAnio(anio));
    if (!sh) return { ok:false, error:'Pestaña ' + _egTabDeAnio(anio) + ' no encontrada' };

    var last = sh.getLastRow();
    if (last < 2) return { ok:true, anio:anio, procesados:0, actualizados:0, yaListos:0,
                           sinArchivo:0, sinXml:0, errores:[], truncado:false, siguiente:0, faltan:0 };

    var cols = _egFacturaCols(sh);                    // crea las columnas si aún no existen
    var iLink = -1;
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                 .map(function (h) { return String(h).trim().toLowerCase(); });
    for (var c = 0; c < hdrs.length; c++) if (hdrs[c].indexOf('link factura') > -1) { iLink = c + 1; break; }
    if (iLink < 0) return { ok:false, error:'No se encontró la columna "Link Factura" en ' + _egTabDeAnio(anio) };

    // Lecturas en BLOQUE (1 llamada por columna, no 1 por celda): el hipervínculo
    // vive en el richText, pero una referencia de texto vive en el valor plano.
    var nFilas = last - 1;
    var richLink = sh.getRange(2, iLink, nFilas, 1).getRichTextValues();
    var valLink  = sh.getRange(2, iLink, nFilas, 1).getValues();
    var valUuid  = sh.getRange(2, cols.uuid, nFilas, 1).getValues();

    function urlDe(i) {
      var r = richLink[i][0];
      var u = (r && r.getLinkUrl()) ? r.getLinkUrl() : String(valLink[i][0] || '');
      return /^https?:\/\//i.test(u) ? u : '';   // una referencia de texto NO es un archivo
    }
    function pendiente(i) {
      if (!urlDe(i)) return false;                                  // sin factura vinculada
      if (!sobrescribir && String(valUuid[i][0] || '').trim()) return false; // ya tiene UUID
      return true;
    }

    var procesados = 0, actualizados = 0, yaListos = 0, sinArchivo = 0, sinXml = 0;
    var errores = [], truncado = false, siguiente = 0;
    var escribir = {};   // rowNum → datos (se vuelcan por columna al final)

    for (var r0 = desde; r0 <= last; r0++) {
      var i = r0 - 2;
      if (!urlDe(i)) { sinArchivo++; continue; }
      if (!pendiente(i)) { yaListos++; continue; }
      // Presupuesto de tiempo: se checa ANTES de abrir el archivo (lo caro).
      if (Date.now() - t0 > maxMs) { truncado = true; siguiente = r0; break; }

      var m = urlDe(i).match(/[-\w]{25,}/);
      if (!m) { sinArchivo++; continue; }
      procesados++;
      var lec = _egLeerFacturaDrive(m[0]);
      if (lec.tipo === 'xml' && lec.datos) { escribir[r0] = lec.datos; actualizados++; }
      else if (lec.tipo === 'otro') sinXml++;       // PDF: no hay UUID que sacar, no se inventa
      else { errores.push('fila ' + r0 + ': ' + (lec.error || 'no se pudo leer')); }
    }

    // Volcado: 4 setValues (uno por columna) sobre el tramo recorrido, en vez de
    // 4 escrituras por fila. Las filas que no tocamos conservan su valor actual.
    var hasta = truncado ? (siguiente - 1) : last;
    if (hasta >= desde && actualizados) {
      var n = hasta - desde + 1;
      var curUuid  = sh.getRange(desde, cols.uuid,  n, 1).getValues();
      var curFolio = sh.getRange(desde, cols.folio, n, 1).getValues();
      var curRfc   = sh.getRange(desde, cols.rfc,   n, 1).getValues();
      var curRazon = sh.getRange(desde, cols.razon, n, 1).getValues();
      Object.keys(escribir).forEach(function (k) {
        var idx = parseInt(k, 10) - desde, d = escribir[k];
        if (idx < 0 || idx >= n) return;
        curUuid[idx][0] = d.uuid; curFolio[idx][0] = d.folio;
        curRfc[idx][0] = d.rfc;   curRazon[idx][0] = d.razon;
      });
      sh.getRange(desde, cols.uuid,  n, 1).setValues(curUuid);
      sh.getRange(desde, cols.folio, n, 1).setValues(curFolio);
      sh.getRange(desde, cols.rfc,   n, 1).setValues(curRfc);
      sh.getRange(desde, cols.razon, n, 1).setValues(curRazon);
    }

    // Cuántas filas quedan pendientes DESPUÉS del corte (para reportar de verdad
    // cuánto falta, no un "se acabó el tiempo" a secas).
    var faltan = 0;
    if (truncado) for (var r1 = siguiente; r1 <= last; r1++) if (pendiente(r1 - 2)) faltan++;

    try {
      logAudit(body.usuario || 'sistema', 'Comprobantes', 'BackfillFacturasEgresos', String(anio),
        'filas ' + desde + '-' + hasta, '', actualizados + ' actualizados' + (truncado ? ' (truncado, faltan ' + faltan + ')' : ''));
    } catch (e) {}

    return { ok:true, anio:anio, desde:desde, hasta:hasta,
             procesados:procesados, actualizados:actualizados, yaListos:yaListos,
             sinArchivo:sinArchivo, sinXml:sinXml, errores:errores,
             truncado:truncado, siguiente:(truncado ? siguiente : 0), faltan:faltan,
             ms:(Date.now() - t0) };
  } catch (ex) { return { ok:false, error: ex.message }; }
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
  ['comprobante-pago', 'Correo de pago (proveedor)',
   'Comprobante de pago — {{empresa}} · {{proveedor}}',
   'Estimado(a) {{proveedor}}:\n\nPor este medio le confirmamos el pago realizado por {{empresa}} correspondiente a:\n\n• Concepto: {{concepto}}\n• Importe: {{monto}}\n• Fecha de pago: {{fecha}}\n• Referencia / póliza: {{poliza}}\n• Folio interno: {{folio}}\n\nAdjuntamos el comprobante de pago para su registro.\n\nQuedamos atentos a cualquier aclaración.\n\nSaludos cordiales,\n{{empresa}} · Administración'],
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

// Núcleo de envío compartido por enviarDocumentoCorreo y enviarCorreoPago.
// El asunto/cuerpo llegan YA interpolados desde el frontend (las variables
// {{...}} se resuelven en la vista); aquí solo se validan, se adjuntan los
// archivos reales de Drive y se manda con GmailApp desde la cuenta del sistema.
function _correoEnviarCore(body) {
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
  // El envío real depende de que los permisos de Gmail estén autorizados en
  // el deployment; si no lo están, GmailApp lanza y se reporta el motivo.
  try {
    GmailApp.sendEmail(para, asunto, cuerpo, opts);
  } catch (se) {
    var m = String(se && se.message || se);
    if (/authoriz|permission|scope|OAuth|acceso|autoriz/i.test(m))
      return { ok: false, error: 'No se pudo enviar: falta autorizar el acceso a Gmail en el deployment. Un administrador debe abrir el proyecto de Apps Script y conceder los permisos de correo. Detalle: ' + m };
    return { ok: false, error: 'No se pudo enviar el correo: ' + m };
  }

  // Bitácora
  try {
    setupPlantillasCorreo();
    var log = SpreadsheetApp.openById(SHEET_ID).getSheetByName(CORREO_LOG_TAB);
    log.appendRow([new Date(), body.usuario || '', para, cc, asunto, body.tipoDoc || '', body.referencia || '', nombres.join(', ')]);
  } catch (le) {}
  try { logAudit(body.usuario || 'sistema', 'Correo', 'Enviar', body.referencia || '', '', '', para + ' | ' + asunto); } catch (ae) {}

  return { ok: true, para: para, adjuntos: nombres.length, fallidos: fallidos.length };
}

function enviarDocumentoCorreo(body) {
  try { return _correoEnviarCore(body); }
  catch (ex) { return { ok: false, error: ex.message }; }
}

// ¿El usuario (por email) puede enviar correos del sistema? Se apoya en la
// hoja Roles (permiso operativo `docs_enviar` o `*`; admin/director siempre).
// Si el módulo de Usuarios no está configurado (modo abierto/dev), se permite.
function _correoUsuarioPuedeEnviar(email) {
  try {
    email = String(email || '').trim();
    var ss = SpreadsheetApp.openById(SHEET_ID);
    if (!ss.getSheetByName('Usuarios')) return true;      // modo abierto
    if (!email) return false;
    var user = getUserRow(ss, email);
    if (!user) return false;
    var rol = String(user.rol || '').toLowerCase();
    if (rol === 'admin' || rol === 'director') return true;
    var perms = getRolConfig(ss, user.rol).permisosOperativos || [];
    if (!perms.length) {   // rol sin permisos configurados → no restringir usuarios existentes
      var restringidos = { socio:1, viewer:1, invitado:1, consulta:1, externo:1 };
      return !restringidos[rol];
    }
    return perms.indexOf('*') !== -1 || perms.indexOf('docs_enviar') !== -1;
  } catch (e) { return true; }   // ante error de config, no bloquear el envío intencional
}

// Correo de PAGO a un proveedor (o interno): mismo motor de envío, pero
// gateado por permiso `docs_enviar` y marcado como 'comprobante-pago' en la
// bitácora. El frontend ya arma asunto/cuerpo con la plantilla de pago y
// adjunta el comprobante/ factura si existen. NUNCA se envía solo: el usuario
// confirma en el diálogo antes de llamar aquí.
function enviarCorreoPago(body) {
  try {
    if (!_correoUsuarioPuedeEnviar(body && body.usuario))
      return { ok: false, error: 'No tienes permiso para enviar correos de pago (docs_enviar).' };
    body = body || {};
    body.tipoDoc = body.tipoDoc || 'comprobante-pago';
    return _correoEnviarCore(body);
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ═════════════════════════════════════════════════════════════════════════
   Envoltorio de UN CLIC para el back-fill fiscal de Egresos (build .209/.210).
   El botón ▶ del editor ejecuta sin argumentos, y backfillFacturasEgresos
   necesita {anio, desde} para reanudarse tras el corte de 6 min de Apps
   Script. Este wrapper encadena las corridas él solo dentro de un
   presupuesto seguro (~4.5 min) y deja en el log un mensaje inequívoco:
   o "TERMINADO", o "vuelve a ejecutar esta misma función" (retoma donde
   quedó, es idempotente — lo ya lleno se salta).
   Ejecutar por año: backfillEgresos2026 / 2025 / 2024.
   ═════════════════════════════════════════════════════════════════════════ */
function _backfillEgresosAnio(anio) {
  var t0 = Date.now(), PRESUPUESTO = 270000;   // 4.5 min < corte de 6 min
  var desde = 2, total = { procesados: 0, yaListos: 0, sinXml: 0 }, r = null;
  do {
    r = backfillFacturasEgresos({ anio: anio, desde: desde, maxMs: 60000 });
    if (!r || !r.ok) { Logger.log('ERROR año ' + anio + ': ' + ((r && r.error) || 'desconocido')); return r; }
    total.procesados += (r.procesados || 0);
    total.yaListos  += (r.yaListos  || 0);
    total.sinXml    += (r.sinXml    || 0);
    desde = r.siguiente || (desde + 1);
    Logger.log('  …fila ' + desde + ' · llenados ' + total.procesados + ' · ya listos ' + total.yaListos + ' · sin XML ' + total.sinXml + (r.truncado ? (' · faltan ' + r.faltan) : ''));
  } while (r.truncado && (Date.now() - t0) < PRESUPUESTO);
  if (r.truncado) {
    Logger.log('AÚN FALTAN ' + r.faltan + ' del año ' + anio + ' → VUELVE A EJECUTAR ESTA MISMA FUNCIÓN (retoma sola donde quedó).');
  } else {
    Logger.log('TERMINADO año ' + anio + ': llenados ' + total.procesados + ', ya estaban ' + total.yaListos + ', sin XML (solo PDF) ' + total.sinXml + '.');
  }
  return { ok: true, anio: anio, terminado: !r.truncado, resumen: total };
}
function backfillEgresos2026() { return _backfillEgresosAnio(2026); }
function backfillEgresos2025() { return _backfillEgresosAnio(2025); }
function backfillEgresos2024() { return _backfillEgresosAnio(2024); }
