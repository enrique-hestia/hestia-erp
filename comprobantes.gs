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

var COMPROBANTES_ROOT_ID = '1rIWggcMKPAtCRvRxBrQgYzSaCK6kp63w';
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
function listComprobantesEstructura() {
  try {
    var root = DriveApp.getFolderById(COMPROBANTES_ROOT_ID);
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

/* ── Leer un mes: XMLs parseados + PDFs emparejados + cruce Egresos ── */
function readComprobantesMes(anio, mes) {
  try {
    anio = parseInt(anio, 10) || new Date().getFullYear();
    mes = parseInt(mes, 10) || (new Date().getMonth() + 1);

    var est = listComprobantesEstructura();
    if (!est.ok) return est;
    var anioObj = null;
    for (var a = 0; a < est.anios.length; a++) if (est.anios[a].anio === anio) { anioObj = est.anios[a]; break; }
    if (!anioObj) return { ok: true, comprobantes: [], estructura: est.anios, aviso: 'No existe carpeta para el año ' + anio };
    var mesObj = null;
    for (var m = 0; m < anioObj.meses.length; m++) if (anioObj.meses[m].mes === mes) { mesObj = anioObj.meses[m]; break; }
    if (!mesObj) return { ok: true, comprobantes: [], estructura: est.anios, aviso: 'No existe carpeta del mes ' + mes + ' en ' + anio };

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
      ok: true, anio: anio, mes: mes, carpeta: mesObj.carpeta,
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

/* ── Vincular: escribe el hipervínculo del XML (y PDF si existe) en la
   columna Link Factura del egreso elegido — igual que el flujo manual ── */
function vincularComprobanteEgreso(body) {
  try {
    var anio = parseInt(body.anio, 10) || new Date().getFullYear();
    var rowNum = parseInt(body.rowNum, 10);
    var fileId = String(body.fileId || '').trim();
    if (!rowNum || !fileId) return { ok: false, error: 'Faltan rowNum o fileId' };

    var ssId = EGRESOS_IDS[anio] || EGRESOS_SS_2026;
    var tabName = (typeof EGRESOS_TABS !== 'undefined' && EGRESOS_TABS[anio]) ? EGRESOS_TABS[anio] : 'Egresos' + anio;
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(tabName) || ss.getSheets()[0];

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var iLinkFact = -1;
    for (var c = 0; c < headers.length; c++) if (headers[c].indexOf('link factura') > -1) { iLinkFact = c; break; }
    if (iLinkFact < 0) return { ok: false, error: 'No se encontró la columna Link Factura en ' + tabName };

    var url = 'https://drive.google.com/file/d/' + fileId + '/view';
    var rich = SpreadsheetApp.newRichTextValue().setText(body.etiqueta || 'Factura XML').setLinkUrl(url).build();
    sheet.getRange(rowNum, iLinkFact + 1).setRichTextValue(rich);

    try { logAudit(body.usuario || 'sistema', 'Comprobantes', 'Vincular', 'fila ' + rowNum, '', '', (body.uuid || fileId)); } catch (e) {}
    return { ok: true, rowNum: rowNum, url: url };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Menú: agrega "Comprobantes" junto a Cuentas por Pagar (mismo
   grupo padre que fin-cxp). Correr UNA VEZ desde el editor. ───────── */
function configurarMenuComprobantes() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Menu');
  if (!sh) return { ok: false, error: 'No se encontró la hoja Menu' };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var idxCol = hdrs.indexOf('ID'), padreCol = hdrs.indexOf('Padre'), ordenCol = hdrs.indexOf('Orden');

  var padre = '', ordenMax = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][idxCol] || '').trim();
    if (id === 'comprobantes') return { ok: true, aviso: 'Ya existe la entrada "comprobantes" en el menú' };
    if (id === 'fin-cxp') padre = String(data[i][padreCol] || '').trim();
  }
  if (padre) {
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][padreCol] || '').trim() === padre) {
        ordenMax = Math.max(ordenMax, parseInt(data[j][ordenCol], 10) || 0);
      }
    }
  }
  var fila = ['comprobantes', padre, 'Comprobantes', '', 'paperclip', ordenMax + 1, 'vista', 'comprobantes', 'TRUE'];
  sh.getRange(sh.getLastRow() + 1, 1, 1, fila.length).setValues([fila]);
  return { ok: true, padre: padre || '(raíz)', orden: ordenMax + 1 };
}
