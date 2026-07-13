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
  'ID', 'Periodo', 'Tipo', 'NumeroDeclaracion', 'FechaPresentacion', 'ImporteDeclarado',
  'IngresosNominales', 'CoeficienteUtilidad', 'UtilidadFiscal', 'PerdidasAplicadas',
  'PTU', 'OtrasDeducciones', 'BaseGravable', 'TasaISR', 'ISRCausado',
  'PagosProvPrevios', 'RetencionesISR', 'OtrosAcreditamientos', 'ISRaCargo',
  'Notas', 'Usuario', 'Actualizado'
];

function _declSheet() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(DECL_TAB);
  if (!sh) {
    sh = ss.insertSheet(DECL_TAB);
    sh.getRange(1, 1, 1, DECL_HEADERS.length).setValues([DECL_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    // Si la hoja existe pero le faltan columnas nuevas, completarlas sin borrar datos.
    var firstRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), DECL_HEADERS.length)).getValues()[0];
    var faltan = false;
    for (var i = 0; i < DECL_HEADERS.length; i++) { if (String(firstRow[i] || '') !== DECL_HEADERS[i]) faltan = true; }
    if (faltan) sh.getRange(1, 1, 1, DECL_HEADERS.length).setValues([DECL_HEADERS]).setFontWeight('bold');
  }
  return sh;
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
    var id = String(d.ID || (periodo + '|' + tipo));

    var sh = _declSheet();
    var data = sh.getDataRange().getValues();
    var h = data[0];
    function idx(name) { return h.indexOf(name); }

    var email = '';
    try { email = verifyToken(body.token || '') || ''; } catch (e) {}
    if (!email) email = String(d.Usuario || '');
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd HH:mm');

    var numCols = { ImporteDeclarado:1, IngresosNominales:1, CoeficienteUtilidad:1, UtilidadFiscal:1,
      PerdidasAplicadas:1, PTU:1, OtrasDeducciones:1, BaseGravable:1, TasaISR:1, ISRCausado:1,
      PagosProvPrevios:1, RetencionesISR:1, OtrosAcreditamientos:1, ISRaCargo:1 };
    var vals = h.map(function (col) {
      col = String(col);
      if (col === 'ID') return id;
      if (col === 'Periodo') return periodo;
      if (col === 'Tipo') return tipo;
      if (col === 'Usuario') return email;
      if (col === 'Actualizado') return stamp;
      var v = d[col];
      if (v === undefined || v === null || v === '') return numCols[col] ? 0 : '';
      if (numCols[col]) { var n = parseFloat(String(v).replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
      return v;
    });

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rid = String(data[i][idx('ID')] || '');
      var rper = String(data[i][idx('Periodo')] || '');
      var rtipo = String(data[i][idx('Tipo')] || '');
      if (rid === id || (rper === periodo && rtipo === tipo)) { foundRow = i + 1; break; }
    }
    if (foundRow > 0) sh.getRange(foundRow, 1, 1, h.length).setValues([vals]);
    else sh.appendRow(vals);

    try { logAudit(email || 'sistema', 'Declaraciones', foundRow > 0 ? 'Editar' : 'Crear', periodo + ' · ' + tipo, '', '', String(d.NumeroDeclaracion || '')); } catch (e) {}
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
