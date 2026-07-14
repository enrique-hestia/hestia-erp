/**
 * declaraciones.gs — Captura y consulta de DECLARACIONES fiscales presentadas.
 *
 * El usuario captura A MANO la declaración que ya presentó ante el SAT (para
 * consultarla después): número de declaración, importe declarado, y "cómo llegó
 * a ese número" — el cálculo del pago provisional de ISR de persona moral
 * (régimen general, Título II LISR, Art. 14):
 *
 *   Ingresos nominales acumulados
 *   × Coeficiente de utilidad            (utilidad fiscal ÷ ingresos nominales del ejercicio anterior)
 *   = Utilidad fiscal del periodo
 *   − Pérdidas fiscales de ejercicios anteriores
 *   − PTU pagada en el ejercicio
 *   − Otras deducciones/anticipos
 *   = Base gravable
 *   × Tasa de ISR (30%)
 *   = ISR causado
 *   − Pagos provisionales efectuados con anterioridad
 *   − Retenciones de ISR (ej. bancos)
 *   − Otros acreditamientos / estímulos
 *   = ISR a cargo  → normalmente = Importe declarado (el SAT puede redondear)
 *
 * Persistencia: hoja "Declaraciones" en el spreadsheet de Ingresos
 * (INGRESOS_SS_ID). Upsert por Periodo (YYYY-MM) + Tipo. Solo lectura para todos;
 * captura/edición requieren permiso 'editar_ingresos'.
 *
 * Rutas: GET  action=declaraciones&periodo=YYYY-MM   (core.gs → readDeclaraciones)
 *        POST action=saveDeclaracion {token, decl}   (finance.gs → saveDeclaracion)
 *        POST action=deleteDeclaracion {token, id}    (finance.gs → deleteDeclaracion)
 */

var DECL_TAB = 'Declaraciones';
var DECL_HEADERS = [
  'ID', 'Periodo', 'Tipo', 'TipoPresentacion', 'Consecutivo', 'NumeroDeclaracion', 'FechaPresentacion', 'ImporteDeclarado',
  'IngresosNominales', 'CoeficienteUtilidad', 'UtilidadFiscal', 'PerdidasAplicadas',
  'PTU', 'OtrasDeducciones', 'BaseGravable', 'TasaISR', 'ISRCausado',
  'PagosProvPrevios', 'RetencionesISR', 'OtrosAcreditamientos', 'ISRaCargo',
  'DeclaracionPdfUrl', 'DeclaracionPdfNombre', 'DeclaracionPdfId',
  'PagoPdfUrl', 'PagoPdfNombre', 'PagoPdfId',
  'Notas', 'Usuario', 'Actualizado'
];

function _declSheet() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(DECL_TAB);
  if (!sh) {
    sh = ss.insertSheet(DECL_TAB);
    sh.getRange(1, 1, 1, DECL_HEADERS.length).setValues([DECL_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // Hoja existente: APPEND de columnas nuevas al final (sin reordenar → no
  // descuadra datos ya capturados; todo se lee/escribe por NOMBRE de columna).
  var lastCol = Math.max(1, sh.getLastColumn());
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v); });
  var toAdd = [];
  for (var i = 0; i < DECL_HEADERS.length; i++) {
    if (existing.indexOf(DECL_HEADERS[i]) === -1) toAdd.push(DECL_HEADERS[i]);
  }
  if (toAdd.length) {
    sh.getRange(1, existing.length + 1, 1, toAdd.length).setValues([toAdd]);
    sh.getRange(1, 1, 1, existing.length + toAdd.length).setFontWeight('bold');
  }
  return sh;
}

// Carpeta de Drive para los PDF adjuntos de declaraciones (por año).
function _declFolder(periodo) {
  var rootIt = DriveApp.getFoldersByName('Declaraciones_Adjuntos');
  var root = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder('Declaraciones_Adjuntos');
  var m = String(periodo || '').match(/^(\d{4})/);
  if (!m) return root;
  var sub = root.getFoldersByName(m[1]);
  return sub.hasNext() ? sub.next() : root.createFolder(m[1]);
}

// Sube un PDF (declaración o comprobante de pago) a Drive y devuelve su URL.
// cual = 'declaracion' | 'pago'. Base64 (data URL o crudo). Máx 12 MB.
function subirDeclaracionPdf(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_ingresos')) {
      return { ok: false, error: 'Sin autorización para adjuntar (editar_ingresos).' };
    }
    var name = String((body && body.filename) || 'declaracion.pdf').trim() || 'declaracion.pdf';
    var mime = String((body && body.mimeType) || 'application/pdf');
    var data = String((body && body.dataBase64) || '').replace(/^data:[^;]+;base64,/, '');
    if (!data) return { ok: false, error: 'Archivo vacío.' };
    var bytes = Utilities.base64Decode(data);
    if (bytes.length > 12 * 1024 * 1024) return { ok: false, error: 'Archivo muy grande (máx 12 MB).' };
    var per = String((body && body.periodo) || '');
    var tipo = String((body && body.tipo) || '');
    var cual = String((body && body.cual) || 'doc');
    var pref = (per ? per + ' ' : '') + (tipo ? tipo + ' ' : '') + cual;
    var blob = Utilities.newBlob(bytes, mime, pref + ' - ' + name);
    var f = _declFolder(per).createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return { ok: true, url: f.getUrl(), fileId: f.getId(), nombre: name };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function readDeclaraciones(periodo) {
  try {
    var sh = _declSheet();
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, items: [] };
    var h = data[0];
    var items = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var o = {};
      for (var c = 0; c < h.length; c++) o[String(h[c])] = row[c];
      if (!o.ID && !o.Periodo) continue;
      if (periodo && String(o.Periodo) !== String(periodo)) continue;
      // Fechas a texto YYYY-MM-DD para el front.
      if (o.FechaPresentacion instanceof Date) {
        o.FechaPresentacion = Utilities.formatDate(o.FechaPresentacion, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd');
      }
      items.push(o);
    }
    // Orden: más reciente primero (por Periodo desc, luego Tipo).
    items.sort(function (a, b) {
      var pa = String(a.Periodo || ''), pb = String(b.Periodo || '');
      if (pa !== pb) return pa < pb ? 1 : -1;
      return String(a.Tipo || '') < String(b.Tipo || '') ? -1 : 1;
    });
    return { ok: true, items: items };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function saveDeclaracion(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_ingresos')) {
      return { ok: false, error: 'Sin autorización para capturar declaraciones. Solicita el permiso editar_ingresos.' };
    }
    var d = body.decl || {};
    var periodo = String(d.Periodo || '').trim();
    var tipo = String(d.Tipo || 'ISR provisional').trim() || 'ISR provisional';
    if (!/^\d{4}-\d{2}$/.test(periodo)) return { ok: false, error: 'Periodo inválido (usa YYYY-MM).' };
    // ID ÚNICO por registro: así una complementaria SIEMPRE es fila nueva y nunca
    // sustituye la normal. Solo al EDITAR (ID presente) se actualiza esa fila.
    var id = String(d.ID || '').trim();
    var esNuevo = !id;
    if (esNuevo) id = 'D' + new Date().getTime() + '-' + Math.floor(Math.random() * 10000);

    var sh = _declSheet();
    var data = sh.getDataRange().getValues();
    var h = data[0];
    function idx(name) { return h.indexOf(name); }

    var email = '';
    try { email = verifyToken(body.token || '') || ''; } catch (e) {}
    if (!email) email = String(d.Usuario || '');
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd HH:mm');

    var numCols = { Consecutivo:1, ImporteDeclarado:1, IngresosNominales:1, CoeficienteUtilidad:1, UtilidadFiscal:1,
      PerdidasAplicadas:1, PTU:1, OtrasDeducciones:1, BaseGravable:1, TasaISR:1, ISRCausado:1,
      PagosProvPrevios:1, RetencionesISR:1, OtrosAcreditamientos:1, ISRaCargo:1 };
    var vals = h.map(function (col) {
      col = String(col);
      if (col === 'ID') return id;
      if (col === 'Periodo') return periodo;
      if (col === 'Tipo') return tipo;
      if (col === 'TipoPresentacion') return String(d.TipoPresentacion || 'Normal') || 'Normal';
      if (col === 'Usuario') return email;
      if (col === 'Actualizado') return stamp;
      var v = d[col];
      if (v === undefined || v === null || v === '') return numCols[col] ? 0 : '';
      if (numCols[col]) { var n = parseFloat(String(v).replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
      return v;
    });

    // Upsert SOLO por ID (nunca por Periodo+Tipo → complementarias conviven).
    var foundRow = -1;
    if (!esNuevo) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idx('ID')] || '') === id) { foundRow = i + 1; break; }
      }
    }
    if (foundRow > 0) sh.getRange(foundRow, 1, 1, h.length).setValues([vals]);
    else sh.appendRow(vals);

    try { logAudit(email || 'sistema', 'Declaraciones', foundRow > 0 ? 'Editar' : 'Crear', periodo + ' · ' + tipo + ' · ' + String(d.TipoPresentacion || 'Normal'), '', '', String(d.NumeroDeclaracion || '')); } catch (e) {}
    return { ok: true, id: id, actualizado: stamp, usuario: email };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function deleteDeclaracion(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_ingresos')) {
      return { ok: false, error: 'Sin autorización para borrar declaraciones.' };
    }
    var id = String(body.id || '');
    if (!id) return { ok: false, error: 'id requerido' };
    var sh = _declSheet();
    var data = sh.getDataRange().getValues();
    var iID = data[0].indexOf('ID');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][iID] || '') === id) {
        sh.deleteRow(i + 1);
        try { logAudit((verifyToken(body.token || '') || 'sistema'), 'Declaraciones', 'Borrar', id, '', '', ''); } catch (e) {}
        return { ok: true };
      }
    }
    return { ok: false, error: 'Declaración no encontrada' };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
