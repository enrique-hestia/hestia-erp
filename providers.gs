/* ==============================================================
   providers.gs — Catálogo centralizado de Proveedores
   --------------------------------------------------------------
   Vive en el spreadsheet de CxP (CXP_SS_ID), pestaña "Proveedores".
   Fuente única para los dropdowns de proveedor en todo el ERP
   (Egresos, Cuentas por Pagar, Cuentas por Cobrar, etc.).
   Vínculo entre módulos = NOMBRE / Razón social (col B).

   Acciones expuestas:
     GET  ?action=proveedores            → readProveedores()
     POST {action:'setupProveedores'}    → setupProveedores()
     POST {action:'saveProveedor', ...}  → saveProveedor(body)
     POST {action:'updateProveedor', ...}→ updateProveedor(body)
   ============================================================== */

var PROV_TAB = 'Proveedores';
var PROV_HEADERS = ['ID','Nombre / Razón social','Nombre comercial','RFC','Categoría',
                    'Contacto','Teléfono','Email','Banco','CLABE / Cuenta','Días crédito',
                    'Estatus','Notas','Fecha alta','Usuario alta'];
var PROV_CATEGORIAS = ['Medicamentos','Laboratorio','Insumos','Servicios','Renta','Honorarios','Nómina','Otros'];

/* ── Crea / formatea la hoja Proveedores (correr una vez) ───────── */
function setupProveedores() {
  try {
    var ss = SpreadsheetApp.openById(CXP_SS_ID);
    var sh = ss.getSheetByName(PROV_TAB);
    if (!sh) sh = ss.insertSheet(PROV_TAB);

    // Encabezados
    sh.getRange(1, 1, 1, PROV_HEADERS.length).setValues([PROV_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a')
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 32);

    // Anchos de columna
    var widths = [90, 230, 170, 120, 130, 160, 110, 200, 130, 170, 95, 95, 240, 110, 150];
    for (var w = 0; w < widths.length; w++) sh.setColumnWidth(w + 1, widths[w]);

    // Validación: Categoría (col E = 5)
    var catRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(PROV_CATEGORIAS, true).setAllowInvalid(true).build();
    sh.getRange(2, 5, sh.getMaxRows() - 1, 1).setDataValidation(catRule);

    // Validación: Estatus (col L = 12)
    var estRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Activo', 'Inactivo'], true).setAllowInvalid(false).build();
    sh.getRange(2, 12, sh.getMaxRows() - 1, 1).setDataValidation(estRule);

    // Banding sutil
    try {
      var rng = sh.getRange(1, 1, sh.getMaxRows(), PROV_HEADERS.length);
      var bandings = sh.getBandings();
      for (var b = 0; b < bandings.length; b++) bandings[b].remove();
      rng.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    } catch (e) {}

    return { ok: true, message: 'Hoja "' + PROV_TAB + '" creada y formateada.', tab: PROV_TAB, columnas: PROV_HEADERS.length };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}

/* ── Lee todos los proveedores + KPIs ──────────────────────────── */
function readProveedores() {
  try {
    var ss = SpreadsheetApp.openById(CXP_SS_ID);
    var sh = ss.getSheetByName(PROV_TAB);
    if (!sh) return { ok: true, rows: [], kpis: { total: 0, activos: 0, inactivos: 0 }, categorias: PROV_CATEGORIAS, _setup: false };
    var raw = sh.getDataRange().getValues();
    if (raw.length < 2) return { ok: true, rows: [], kpis: { total: 0, activos: 0, inactivos: 0 }, categorias: PROV_CATEGORIAS, _setup: true };

    function dt(v) { if (!v) return ''; if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0'); return String(v); }
    var rows = [], activos = 0, inactivos = 0, catSet = {};

    for (var i = 1; i < raw.length; i++) {
      var r = raw[i];
      var nombre = String(r[1] || '').trim();
      if (!nombre) continue; // fila vacía
      var estatus = String(r[11] || 'Activo').trim() || 'Activo';
      if (estatus.toLowerCase() === 'inactivo') inactivos++; else activos++;
      var categoria = String(r[4] || '').trim();
      if (categoria) catSet[categoria] = 1;
      rows.push({
        _rowNum: i + 1,
        id: String(r[0] || '').trim(),
        nombre: nombre,
        comercial: String(r[2] || '').trim(),
        rfc: String(r[3] || '').trim(),
        categoria: categoria,
        contacto: String(r[5] || '').trim(),
        telefono: String(r[6] || '').trim(),
        email: String(r[7] || '').trim(),
        banco: String(r[8] || '').trim(),
        cuenta: String(r[9] || '').trim(),
        diasCredito: r[10] === '' || r[10] == null ? '' : Number(r[10]) || 0,
        estatus: estatus,
        notas: String(r[12] || '').trim(),
        fechaAlta: dt(r[13]),
        usuarioAlta: String(r[14] || '').trim()
      });
    }
    rows.sort(function (a, b) { return a.nombre.toLowerCase() < b.nombre.toLowerCase() ? -1 : 1; });
    var categorias = Object.keys(catSet);
    PROV_CATEGORIAS.forEach(function (c) { if (categorias.indexOf(c) === -1) categorias.push(c); });

    return { ok: true, rows: rows, kpis: { total: rows.length, activos: activos, inactivos: inactivos }, categorias: categorias, _setup: true };
  } catch (ex) {
    return { ok: false, error: ex.message, rows: [] };
  }
}

/* ── Siguiente ID PROV-##### ────────────────────────────────────── */
function _getNextProvID(sh) {
  var lr = sh.getLastRow();
  var max = 0;
  if (lr > 1) {
    var ids = sh.getRange(2, 1, lr - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var m = String(ids[i][0] || '').match(/PROV-(\d+)/);
      if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
    }
  }
  return 'PROV-' + String(max + 1).padStart(5, '0');
}

/* ── Mapa nombre→fila para validar duplicados ──────────────────── */
function _provFields(body) {
  return [
    String(body.nombre || '').trim(),
    String(body.comercial || '').trim(),
    String(body.rfc || '').trim(),
    String(body.categoria || '').trim(),
    String(body.contacto || '').trim(),
    String(body.telefono || '').trim(),
    String(body.email || '').trim(),
    String(body.banco || '').trim(),
    String(body.cuenta || '').trim(),
    body.diasCredito === '' || body.diasCredito == null ? '' : (Number(body.diasCredito) || 0),
    String(body.estatus || 'Activo').trim() || 'Activo',
    String(body.notas || '').trim()
  ];
}

/* ── Alta de proveedor ──────────────────────────────────────────── */
function saveProveedor(body) {
  try {
    if (!body || !String(body.nombre || '').trim()) return { ok: false, error: 'El nombre / razón social es obligatorio.' };
    var ss = SpreadsheetApp.openById(CXP_SS_ID);
    var sh = ss.getSheetByName(PROV_TAB);
    if (!sh) { setupProveedores(); sh = ss.getSheetByName(PROV_TAB); }

    // Evitar duplicado por nombre (case-insensitive)
    var lr = sh.getLastRow();
    if (lr > 1) {
      var names = sh.getRange(2, 2, lr - 1, 1).getValues();
      var nuevo = String(body.nombre).trim().toLowerCase();
      for (var i = 0; i < names.length; i++) {
        if (String(names[i][0] || '').trim().toLowerCase() === nuevo) return { ok: false, error: 'Ya existe un proveedor con ese nombre.' };
      }
    }

    var id = _getNextProvID(sh);
    var f = _provFields(body);
    var fecha = new Date();
    var fechaStr = fecha.getFullYear() + '-' + String(fecha.getMonth() + 1).padStart(2, '0') + '-' + String(fecha.getDate()).padStart(2, '0');
    var row = [id].concat(f).concat([fechaStr, String(body.usuario || '')]);
    sh.appendRow(row);
    try { logAudit(body.usuario || '', 'Proveedores', 'Alta', id, 'nombre', '', body.nombre); } catch (e) {}
    return { ok: true, id: id, message: 'Proveedor registrado.' };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}

/* ── Edición de proveedor ───────────────────────────────────────── */
function updateProveedor(body) {
  try {
    if (!body || !body.rowNum) return { ok: false, error: 'Fila no especificada.' };
    if (!String(body.nombre || '').trim()) return { ok: false, error: 'El nombre / razón social es obligatorio.' };
    var ss = SpreadsheetApp.openById(CXP_SS_ID);
    var sh = ss.getSheetByName(PROV_TAB);
    if (!sh) return { ok: false, error: 'La hoja Proveedores no existe.' };
    var rowNum = parseInt(body.rowNum, 10);
    if (rowNum < 2 || rowNum > sh.getLastRow()) return { ok: false, error: 'Fila fuera de rango.' };

    var f = _provFields(body);
    // Cols B..M (2..13) = nombre..notas. ID, fecha alta y usuario alta se conservan.
    sh.getRange(rowNum, 2, 1, f.length).setValues([f]);
    try { logAudit(body.usuario || '', 'Proveedores', 'Edición', String(body.id || rowNum), 'nombre', '', body.nombre); } catch (e) {}
    return { ok: true, message: 'Proveedor actualizado.' };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}
