var SHEET_ID      = '1FMB2Qmv5z36sUDlVpwzjihNzrfS55k8MG32J04IBaR4';
var API_VERSION   = 'v2026-06-08-U';
var AUTH_SECRET   = 'hestia2026erp-secret'; // Cambia esto por algo único

/* ── Autenticación: helpers ──────────────────────────────────── */
function sha256Hex(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str)
    .map(function(b){ return ('0'+(b&0xFF).toString(16)).slice(-2); }).join('');
}
function generateToken(email, day) {
  day = day || fmtDate(new Date());
  return Utilities.base64Encode(email+'|'+day+'|'+sha256Hex(email+'|'+day+'|'+AUTH_SECRET));
}
function verifyToken(token) {
  if (!token) return null;
  try {
    var dec   = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    var parts = dec.split('|');
    if (parts.length < 3) return null;
    var email = parts[0], day = parts[1];
    var diff  = (new Date() - new Date(day)) / 86400000;
    if (diff > 7 || diff < -1) return null;
    return (generateToken(email, day) === token) ? email : null;
  } catch(ex){ return null; }
}
function getUserRow(ss, email) {
  var sh = ss.getSheetByName('Usuarios');
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  var h    = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
  var eI=h.indexOf('email'), nI=h.indexOf('nombre'), pI=h.indexOf('contraseña'),
      rI=h.indexOf('rol'),   aI=h.indexOf('activo');
  for (var i=1; i<data.length; i++) {
    if (String(data[i][eI]).trim().toLowerCase() === email.toLowerCase()) {
      return { email: String(data[i][eI]).trim(), nombre: String(data[i][nI>-1?nI:0]).trim(),
               password: pI>-1?String(data[i][pI]).trim():'',
               rol: rI>-1?String(data[i][rI]).trim():'viewer',
               activo: aI>-1?data[i][aI]:true, rowNum: i+1 };
    }
  }
  return null;
}
function getRolConfig(ss, rol) {
  var sh  = ss.getSheetByName('Roles');
  var def = { vistasBloqueadas:[], soloLectura:false };
  if (!sh) return def;
  var data = sh.getDataRange().getValues();
  var h    = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
  var rI=h.indexOf('rol'), bI=h.indexOf('vistas_bloqueadas'), lI=h.indexOf('solo_lectura');
  for (var i=1; i<data.length; i++) {
    if (String(data[i][rI]).trim().toLowerCase() === rol.toLowerCase()) {
      var bloq = bI>-1 ? String(data[i][bI]).split(',').map(function(v){return v.trim();}).filter(Boolean) : [];
      return { vistasBloqueadas: bloq, soloLectura: lI>-1 ? !!data[i][lI] : false };
    }
  }
  return def;
}

// Mapeo: nombre de pestaña → ID del spreadsheet externo donde se lee/escribe
// Agregar aquí cualquier hoja de captura futura
var LAB_SS_ID = '1hYmIl4gSTVrvghP7KY0y0dC200o8w0zShXj63zP-TrQ';
var MED_SS_ID = '1fiuUtw-sg2ELNxq9bCjaOtRz1n87wuVi8IOQYzEi8tM';
// ⚠️ Reemplaza con el ID real del spreadsheet de Quirofano cuando lo crees
var QX_SS_ID  = LAB_SS_ID;  // placeholder — apunta al lab por ahora

var CAPTURA_SHEETS = {
  // ── Medicamentos ─────────────────────────────────────────────
  'Medicamentos':    MED_SS_ID,
  'Orden_Compra':    MED_SS_ID,
  'Ent. Med':        MED_SS_ID,
  'Lista Med':       MED_SS_ID,
  'Estimulacion':    MED_SS_ID,
  'Estimulación':    MED_SS_ID,
  'Salidas Med':     MED_SS_ID,
  // ── Laboratorio ──────────────────────────────────────────────
  'Resumen':         LAB_SS_ID,
  'ART Lab':         LAB_SS_ID,
  'FET':             LAB_SS_ID,
  'Andrología':      LAB_SS_ID,
  'Andrologia':      LAB_SS_ID,
  'Inventario Crío': LAB_SS_ID,
  'Inventario Crio': LAB_SS_ID,
  'Insumos':         LAB_SS_ID,
  'Salidas Lab':     LAB_SS_ID,
  // ── Quirofano ────────────────────────────────────────────────
  'Insumos Qx':      QX_SS_ID,
  'Salidas Qx':      QX_SS_ID,
  // ── Otras hojas ──────────────────────────────────────────────
  'Pacientes':       '1uoQU-vbefxWwaLxJyTFT25gj7Nr2223WISa3tqH-Rio',
  'Productos':       '1eXskEMPdwuwEuV7GmVDNfyO1ulxhsZ9F_2hDVRDdIAY'
};
// Aliases: nombre alternativo del menú → nombre exacto de la pestaña en Sheets
var SHEET_ALIASES = {
  'Orden_Compra':  'Ent. Med',
  'Estimulacion':  'Estimulación'
};
// Fallback si la hoja no está en el mapeo (usa el sheet principal de Hestia ERP)
var CAPTURA_SHEET_ID_DEFAULT = SHEET_ID;

/* ══════════════════════════════════════════════════════════════
   doGet — Enrutador principal
   ?action=menu                              → menú (carga inicial)
   ?action=view&view=X&fechaInicio=Y&fechaFin=Z → datos de la vista
   ?action=insert&sheet=X&...campos          → inserta fila
   ══════════════════════════════════════════════════════════════ */
function doGet(e) {
  try {
    var ss     = SpreadsheetApp.openById(SHEET_ID);
    var action = (e && e.parameter.action) || 'menu';
    var view   = (e && e.parameter.view)   || 'resumen';

    // tabs: muestra pestañas de cualquier spreadsheet (sin auth, solo para debug)
    if (action === 'tabs') {
      var sid = (e && e.parameter.sid) || CAPTURA_SHEETS['Medicamentos'];
      try {
        var ssTabs = SpreadsheetApp.openById(sid);
        var tabs   = ssTabs.getSheets().map(function(s){
          return { name: s.getName(), gid: s.getSheetId() };
        });
        return jsonResponse({ spreadsheetId: sid, tabs: tabs });
      } catch(ex) { return jsonResponse({ error: ex.message }); }
    }

    // labinspect: lee las primeras filas del spreadsheet de Lab (sin auth, solo debug)
    if (action === 'labinspect') {
      try {
        var ssLab   = SpreadsheetApp.openById('1hYmIl4gSTVrvghP7KY0y0dC200o8w0zShXj63zP-TrQ');
        var allShts = ssLab.getSheets();
        var result  = allShts.map(function(sh) {
          var preview = sh.getRange(1, 1, Math.min(3, sh.getLastRow()), Math.min(10, sh.getLastColumn())).getValues();
          return { name: sh.getName(), gid: sh.getSheetId(), preview: preview };
        });
        return jsonResponse({ sheets: result });
      } catch(ex) { return jsonResponse({ error: ex.message }); }
    }

    // Rango de fechas — default: últimos 6 meses
    var hoy        = new Date();
    var defInicio  = new Date(hoy.getFullYear(), hoy.getMonth() - 5, 1);
    var defFin     = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);
    var fechaInicio = (e && e.parameter.fechaInicio) || fmtDate(defInicio);
    var fechaFin    = (e && e.parameter.fechaFin)    || fmtDate(defFin);

    // ── LOGIN: valida credenciales y devuelve token + permisos ──
    if (action === 'login') {
      var email    = (e && e.parameter.email)    || '';
      var password = (e && e.parameter.password) || '';
      if (!email) return jsonResponse({ error: 'Email requerido.' });
      var shU = ss.getSheetByName('Usuarios');
      if (!shU) return jsonResponse({ error: 'Módulo de usuarios no configurado.' });
      var user = getUserRow(ss, email);
      if (!user)        return jsonResponse({ error: 'Usuario no encontrado.' });
      if (!user.activo) return jsonResponse({ error: 'Usuario inactivo. Contacta al administrador.' });
      if (user.password && user.password !== password)
                        return jsonResponse({ error: 'Contraseña incorrecta.' });
      var rolCfg = getRolConfig(ss, user.rol);
      return jsonResponse({
        success: true, token: generateToken(user.email),
        email: user.email, nombre: user.nombre, rol: user.rol,
        vistasBloqueadas: rolCfg.vistasBloqueadas,
        soloLectura:      rolCfg.soloLectura
      });
    }

    // ── VALIDAR TOKEN en todas las acciones (si la hoja Usuarios existe) ──
    var currentUser = null;
    var shUsuariosExiste = !!ss.getSheetByName('Usuarios');
    if (shUsuariosExiste) {
      var tkn = (e && e.parameter.token) || '';
      var tkEmail = verifyToken(tkn);
      if (!tkEmail) return jsonResponse({ error: 'Sesión inválida. Inicia sesión nuevamente.', code: 401 });
      currentUser = getUserRow(ss, tkEmail);
      if (!currentUser || !currentUser.activo) return jsonResponse({ error: 'Usuario no autorizado.', code: 403 });
    }

    // ── USUARIOS: listado para admin ──
    if (action === 'usuarios') {
      if (!currentUser || currentUser.rol !== 'admin')
        return jsonResponse({ error: 'Sin permisos de administrador.' });
      var shU2 = ss.getSheetByName('Usuarios');
      var rowsU = shU2.getDataRange().getValues();
      var hdrsU = rowsU[0];
      var pIdx  = hdrsU.map(function(h){ return String(h).toLowerCase(); }).indexOf('contraseña');
      var usuarios = rowsU.slice(1).map(function(r, i) {
        var obj = { _rowNum: i+2 };
        hdrsU.forEach(function(h, j) {
          if (j !== pIdx) obj[String(h).trim()] = r[j]; // no devolver contraseña
        });
        return obj;
      }).filter(function(u){ return u['Email'] || u['email']; });
      var shRol = ss.getSheetByName('Roles');
      var roles = [];
      if (shRol) {
        var rowsR = shRol.getDataRange().getValues();
        roles = rowsR.slice(1).map(function(r){ return String(r[0]).trim(); }).filter(Boolean);
      }
      return jsonResponse({ usuarios: usuarios, roles: roles });
    }

    // ── SAVEUSER: crear/actualizar usuario (solo admin) ──
    if (action === 'saveuser') {
      if (!currentUser || currentUser.rol !== 'admin')
        return jsonResponse({ error: 'Sin permisos de administrador.' });
      var shU3   = ss.getSheetByName('Usuarios');
      var hdrs3  = shU3.getRange(1,1,1,shU3.getLastColumn()).getValues()[0];
      var rowNum3= parseInt((e && e.parameter.rowNum) || '0');
      var newRow = hdrs3.map(function(h) {
        var key = String(h).trim();
        return (e.parameter[key] !== undefined) ? e.parameter[key] : '';
      });
      if (rowNum3 > 1) {
        shU3.getRange(rowNum3, 1, 1, hdrs3.length).setValues([newRow]);
      } else {
        shU3.appendRow(newRow);
      }
      return jsonResponse({ success: true });
    }

    // ── SAVEROLE: crea o actualiza un rol (solo admin) ──
    if (action === 'saveRole') {
      if (!currentUser || currentUser.rol.toLowerCase() !== 'admin')
        return jsonResponse({ error: 'Sin permisos de administrador.' });
      var rolName    = (e && e.parameter.rol)              || '';
      var bloqueadas = (e && e.parameter.vistasBloqueadas) || '';
      var soloLec    = (e && e.parameter.soloLectura)      === 'true';
      var desc       = (e && e.parameter.descripcion)      || '';
      if (!rolName) return jsonResponse({ error: 'Nombre de rol requerido.' });
      var shR  = ss.getSheetByName('Roles');
      if (!shR) return jsonResponse({ error: 'Hoja Roles no encontrada.' });
      var rData = shR.getDataRange().getValues();
      var rH    = rData[0].map(function(c){ return String(c).trim().toLowerCase(); });
      var rI  = rH.indexOf('rol');
      var bI  = rH.indexOf('vistas_bloqueadas');
      var lI  = rH.indexOf('solo_lectura');
      var dI  = rH.indexOf('descripcion');
      var rowIdx = -1;
      for (var ri = 1; ri < rData.length; ri++) {
        if (String(rData[ri][rI]).trim().toLowerCase() === rolName.toLowerCase()) { rowIdx = ri + 1; break; }
      }
      if (rowIdx > 0) {
        if (bI > -1) shR.getRange(rowIdx, bI+1).setValue(bloqueadas);
        if (lI > -1) shR.getRange(rowIdx, lI+1).setValue(soloLec);
        if (dI > -1) shR.getRange(rowIdx, dI+1).setValue(desc);
      } else {
        var newRow = Array(rData[0].length).fill('');
        if (rI > -1) newRow[rI] = rolName;
        if (bI > -1) newRow[bI] = bloqueadas;
        if (lI > -1) newRow[lI] = soloLec;
        if (dI > -1) newRow[dI] = desc;
        shR.appendRow(newRow);
      }
      return jsonResponse({ success: true });
    }

    if (action === 'menu') {
      return jsonResponse({
        menu:        readMenu(ss),
        fechaInicio: fechaInicio,
        fechaFin:    fechaFin,
        version:     API_VERSION
      });
    }

    if (action === 'insert') {
      return jsonResponse(insertRow(ss, e));
    }

    // Debug: ?action=debug&sheet=NombreHoja
    if (action === 'debug') {
      var sheetName = (e && e.parameter.sheet) || '';
      var capturaId = getCapturaId(sheetName);
      try {
        var ssDeb = SpreadsheetApp.openById(capturaId);
        var tabs  = ssDeb.getSheets().map(function(s){ return s.getName(); });
        return jsonResponse({ sheetName: sheetName, spreadsheetId: capturaId,
                              tabsFound: tabs, capturaSheets: CAPTURA_SHEETS });
      } catch(ex) {
        return jsonResponse({ error: ex.message, sheetName: sheetName, spreadsheetId: capturaId });
      }
    }

    // options: ?action=options&sheet=NombreHoja → lee validaciones dropdown de la hoja
    if (action === 'options') {
      var sheetName = (e && e.parameter.sheet) || 'Pacientes';
      var capturaId = getCapturaId(sheetName);
      try {
        var ssOpt  = SpreadsheetApp.openById(capturaId);
        var shOpt  = ssOpt.getSheetByName(sheetName);
        if (!shOpt) return jsonResponse({ error: 'Hoja no encontrada: ' + sheetName });
        var lastCol  = shOpt.getLastColumn();
        var headers  = shOpt.getRange(1, 1, 1, lastCol).getValues()[0];
        var targets  = ['Origen', 'Canal', 'Médico Tratante', 'País'];
        var options  = {};
        // Leer opciones desde pestaña "Opciones" (una columna por campo dropdown)
        var shOpciones = ssOpt.getSheetByName('Opciones');
        if (shOpciones) {
          var optHeaders = shOpciones.getRange(1, 1, 1, shOpciones.getLastColumn()).getValues()[0];
          var optData    = shOpciones.getRange(2, 1, Math.max(shOpciones.getLastRow() - 1, 1), shOpciones.getLastColumn()).getValues();
          targets.forEach(function(colName) {
            var colIdx = -1;
            for (var i = 0; i < optHeaders.length; i++) {
              if (String(optHeaders[i]).trim() === colName) { colIdx = i; break; }
            }
            if (colIdx === -1) { options[colName] = []; return; }
            options[colName] = optData
              .map(function(row) { return String(row[colIdx]).trim(); })
              .filter(function(v) { return v && v !== '' && v !== 'undefined'; });
          });
        }
        return jsonResponse({ options: options });
      } catch(ex) {
        return jsonResponse({ error: ex.message });
      }
    }

    // listamedoptions: ?action=listamedoptions → catálogo de medicamentos con costo
    // Lee "Lista Med" del spreadsheet de Medicamentos y devuelve [{nombre, costo}]
    if (action === 'listamedoptions') {
      try {
        var ssLm   = SpreadsheetApp.openById(CAPTURA_SHEETS['Lista Med']);
        var shLm   = ssLm.getSheetByName('Lista Med');
        if (!shLm) return jsonResponse({ error: 'Hoja Lista Med no encontrada.' });
        var lmData = shLm.getDataRange().getValues();
        // Fila 0 = encabezados, col 0 = nombre, col 1 = costo
        var meds = lmData.slice(1)
          .filter(function(r) { return String(r[0]).trim() !== ''; })
          .map(function(r) {
            var costoRaw = String(r[1]).replace(/[$,\s]/g, '');
            return { nombre: String(r[0]).trim(), costo: parseFloat(costoRaw) || 0 };
          });
        return jsonResponse({ meds: meds });
      } catch(ex) {
        return jsonResponse({ error: ex.message });
      }
    }

    // nextid: ?action=nextid&sheet=Pacientes&prefix=HEC → siguiente ID disponible
    if (action === 'nextid') {
      var sheetName = (e && e.parameter.sheet)  || 'Pacientes';
      var prefix    = (e && e.parameter.prefix) || 'HEC';
      var capturaId = getCapturaId(sheetName);
      try {
        var ssNid = SpreadsheetApp.openById(capturaId);
        var shNid = ssNid.getSheetByName(sheetName);
        if (!shNid || shNid.getLastRow() < 2) {
          return jsonResponse({ nextId: prefix + '-001' });
        }
        var ids = shNid.getRange(2, 1, shNid.getLastRow() - 1, 1).getValues()
          .map(function(r) { return String(r[0]); });
        var maxNum = 0;
        ids.forEach(function(id) {
          var m = id.match(/(\d+)$/);
          if (m) maxNum = Math.max(maxNum, parseInt(m[1]));
        });
        return jsonResponse({ nextId: prefix + '-' + String(maxNum + 1).padStart(3, '0') });
      } catch(ex) {
        return jsonResponse({ nextId: prefix + '-001', error: ex.message });
      }
    }

    // update: ?action=update&sheet=X&rowNum=N&Campo=valor → actualiza fila en Sheets
    if (action === 'update') {
      var sheetName = (e && e.parameter.sheet)  || '';
      if (SHEET_ALIASES[sheetName]) sheetName = SHEET_ALIASES[sheetName];
      var rowNum    = parseInt((e && e.parameter.rowNum) || '0');
      if (!sheetName || !rowNum) return jsonResponse({ error: 'sheet y rowNum son requeridos' });
      var capturaId = getCapturaId(sheetName);
      try {
        var ssUpd = SpreadsheetApp.openById(capturaId);
        var shUpd = findSheet(ssUpd, sheetName);
        if (!shUpd) return jsonResponse({ error: 'Hoja no encontrada: ' + sheetName });
        var hdrInfo = getSheetHeaders(shUpd);
        var hdrs = hdrInfo.headers;
        var cur  = shUpd.getRange(rowNum, 1, 1, hdrs.length).getValues()[0];
        var newRow = hdrs.map(function(h, i) {
          return (e.parameter[h] !== undefined) ? e.parameter[h] : cur[i];
        });
        shUpd.getRange(rowNum, 1, 1, hdrs.length).setValues([newRow]);
        invalidateViewCache(sheetName);
        return jsonResponse({ success: true, rowNum: rowNum });
      } catch(ex) {
        return jsonResponse({ error: ex.message });
      }
    }

    // action === 'view'  — con caché de 60 s para reducir lecturas a Sheets
    if (action === 'view') {
      var cache    = CacheService.getScriptCache();
      var cacheKey = 'v2_' + view + '_' + fechaInicio + '_' + fechaFin;
      var cached   = cache.get(cacheKey);
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
      var result   = readViewData(ss, view, fechaInicio, fechaFin);
      var json     = JSON.stringify(result);
      try { cache.put(cacheKey, json, 60); } catch(ignored) {}
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }
    return jsonResponse(readViewData(ss, view, fechaInicio, fechaFin));

  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

/* Formatea Date como YYYY-MM-DD */
function fmtDate(d) {
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  return d.getFullYear() + '-' + mm + '-' + dd;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Borra todas las entradas de caché relacionadas con una hoja (fuente) */
function invalidateViewCache(sheetName) {
  try {
    var cache = CacheService.getScriptCache();
    // La clave incluye el viewId que puede ser el nombre de hoja directamente.
    // Borramos con el nombre exacto y variantes comunes.
    var keys = [];
    // Generar keys para los últimos 12 meses como rango de fechas posible
    var now = new Date();
    for (var i = -1; i <= 12; i++) {
      var d1 = new Date(now.getFullYear(), now.getMonth() - 6 + i, 1);
      var d2 = new Date(now.getFullYear(), now.getMonth() - 6 + i + 7, 0);
      keys.push('v2_' + sheetName + '_' + fmtDate(d1) + '_' + fmtDate(d2));
    }
    cache.removeAll(keys);
  } catch(ignored) {}
}

/* ── getDefaultPeriodo mantenido por compatibilidad (ya no se usa) ─── */
function getDefaultPeriodo(ss) { return ''; }

/* ══════════════════════════════════════════════════════════════
   LEE LA HOJA Menu
   Columnas: ID | Padre | Label | Seccion | Icono | Orden | Tipo | Fuente | Activo
   ══════════════════════════════════════════════════════════════ */
function readMenu(ss) {
  var rows = ss.getSheetByName('Menu').getDataRange().getValues();
  return rows.slice(1)
    .filter(function(r) {
      // Solo filas activas (columna I = TRUE o vacío)
      var activo = r[8];
      return activo !== false && String(activo).toUpperCase() !== 'FALSE';
    })
    .map(function(r) {
      return {
        id:      String(r[0]).trim(),
        padre:   String(r[1]).trim(),
        label:   String(r[2]).trim(),
        seccion: String(r[3]).trim(),
        icono:   String(r[4]).trim() || 'circle',
        orden:   Number(r[5]) || 0,
        tipo:    String(r[6]).trim().toLowerCase(), // 'vista' | 'grupo'
        fuente:  String(r[7]).trim(),               // 'mensual' | NombreHoja
        activo:  r[8] !== false
      };
    });
}

/* ══════════════════════════════════════════════════════════════
   ENRUTADOR DE DATOS POR VISTA
   fuente = 'mensual' → datos financieros filtrados por rango de fechas
   fuente = NombreHoja → lee esa hoja como tabla de captura (sin filtro)
   ══════════════════════════════════════════════════════════════ */
function readViewData(ss, viewId, fechaInicio, fechaFin) {
  var menu = readMenu(ss);
  var item = null;
  for (var i = 0; i < menu.length; i++) {
    if (menu[i].id === viewId) { item = menu[i]; break; }
  }
  var fuente = item ? item.fuente : 'mensual';
  // Traducir alias (ej: Orden_Compra → Ent. Med)
  if (fuente && SHEET_ALIASES[fuente]) fuente = SHEET_ALIASES[fuente];

  // Vistas de configuración local — no leen Sheets
  if (/^(configuracion|ajustes|config)$/i.test(viewId) || /^(configuracion|ajustes|config)$/i.test(fuente||'')) {
    return { viewId: viewId, fuente: fuente, rows: [], headers: [], _isConfig: true };
  }

  if (!fuente || fuente === 'mensual') {
    return readMensualData(ss, fechaInicio, fechaFin, viewId);
  }

  if (fuente === 'med-dashboard' || viewId === 'med-resumen' || viewId === 'prod-medicamentos') {
    return readMedDashboard(ss, fechaInicio, fechaFin);
  }

  if (fuente === 'lab-resumen' || viewId === 'lab-resumen' || viewId === 'lab-resumen-dash') {
    return readLabResumen(fechaInicio, fechaFin);
  }

  if (fuente === 'qx-resumen' || viewId === 'qx-resumen') {
    return readQxResumen(fechaInicio, fechaFin);
  }

  if (viewId === 'gestion-roles') {
    return readGestionRoles(ss);
  }

  if (fuente === 'Rep Ejecutivo' || viewId === 'rep-ejecutivo') {
    return readRepEjecutivo(ss, fechaInicio, fechaFin);
  }

  // Vistas de Reportes: devolver placeholder vacío hasta que se creen las hojas
  if (/^rep-/.test(viewId)) {
    return { view: viewId, fuente: fuente, rows: [], headers: [], periodo: fechaInicio + ' — ' + fechaFin };
  }

  return readCapturaData(ss, fuente, viewId, fechaInicio, fechaFin);
}

/* ══════════════════════════════════════════════════════════════
   RESUMEN LABORATORIO — Dashboard clínico de calidad embrionaria
   Fuente: spreadsheet Lab (LAB_SS_ID)
   Hojas leídas: Resumen, ART Lab, FET, Inventario Crío, Insumos
   ══════════════════════════════════════════════════════════════ */
function readLabResumen(fechaInicio, fechaFin) {

  // ── CONFIGURACIÓN DE COLUMNAS ──────────────────────────────────
  // ART Lab: encabezados en fila 1 + fila 2 fusionadas (A=0, B=1…)
  // Ajusta estos índices si agregas/mueves columnas en la hoja
  var ART_COL_MES       = 0;   // A: Mes-Año
  var ART_COL_FECHA     = 1;   // B: Date
  var ART_COL_OOCITOS   = 9;   // J: # oocytes (recuperados)
  // Cuando conozcas las columnas de 2PN y blastocistos, actualiza:
  var ART_COL_2PN       = -1;  // -1 = no configurado aún
  var ART_COL_BLASTO    = -1;  // -1 = no configurado aún
  var ART_COL_ICSI_DANO = -1;  // -1 = no configurado aún

  // FET: encabezado en fila 2 (fila 1 vacía)
  var FET_COL_FECHA    = 0;  // A: Fecha
  var FET_COL_SURVIVED = 5;  // F: Survived ("Si"/"No")
  var FET_COL_BETA     = 6;  // G: Beta
  var FET_COL_PREG     = 7;  // H: Clinical Preg.

  // Inventario Crío: encabezado en fila 1
  var CRIO_COL_NOMBRE  = 0;  // A: Nombre paciente
  var CRIO_COL_FECHA   = 1;  // B: Fecha Crío
  var CRIO_COL_OOV     = 2;  // C: Oov (ovocitos)
  var CRIO_COL_EMB     = 3;  // D: Emb (embriones)

  // Insumos: encabezado en fila 2 (fila 1 vacía); col A vacía → datos desde col B
  var INS_COL_INSUMO   = 3;  // D: Insumo
  var INS_COL_PROV     = 4;  // E: Proveedor
  var INS_COL_FECHA    = 2;  // C: Fecha
  var INS_COL_COSTO    = 8;  // I: Costo
  // ─────────────────────────────────────────────────────────────

  function pct(num, den) {
    if (!den || den === 0) return null;
    return Math.round((num / den) * 1000) / 10;
  }
  function fmtFecha(v) {
    if (!v) return '';
    if (v instanceof Date) return fmtDate(v);
    var d = new Date(v);
    return isNaN(d.getTime()) ? String(v) : fmtDate(d);
  }
  function mesLabel(v) {
    if (!v) return '';
    if (v instanceof Date) return (v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0'));
    var s = String(v);
    var m = s.match(/(\d{4})-(\d{2})/);
    return m ? m[0] : s.slice(0,7);
  }

  try {
    var ssLab = SpreadsheetApp.openById(LAB_SS_ID);

    // ── 1. Leer ART Lab (encabezados fusionados fila 1+2, datos desde fila 3) ──
    var shArt  = findSheet(ssLab, 'ART Lab');
    var artRaw = shArt ? shArt.getDataRange().getValues() : [];
    // Fusionar headers de fila 1 y fila 2
    var artData = artRaw.slice(2).filter(function(r) {
      return r.some(function(c) { return String(c).trim() !== ''; });
    });
    var totalCiclos = artData.length;
    var totalOocitos = artData.reduce(function(s,r) { return s + (Number(r[ART_COL_OOCITOS])||0); }, 0);

    // % Fecundación y Blastocistos: calcular si las columnas están configuradas
    var total2PN   = ART_COL_2PN   > -1 ? artData.reduce(function(s,r){ return s+(Number(r[ART_COL_2PN])||0); },0) : null;
    var totalBlasto= ART_COL_BLASTO > -1 ? artData.reduce(function(s,r){ return s+(Number(r[ART_COL_BLASTO])||0); },0) : null;

    // Tendencia por mes desde ART Lab
    var artPorMes = {};
    artData.forEach(function(r) {
      var mes = mesLabel(r[ART_COL_FECHA] || r[ART_COL_MES]);
      if (!mes) return;
      if (!artPorMes[mes]) artPorMes[mes] = { ciclos:0, oocitos:0, pn2:0, blasto:0 };
      artPorMes[mes].ciclos++;
      artPorMes[mes].oocitos += Number(r[ART_COL_OOCITOS]) || 0;
      if (ART_COL_2PN    > -1) artPorMes[mes].pn2    += Number(r[ART_COL_2PN])    || 0;
      if (ART_COL_BLASTO > -1) artPorMes[mes].blasto += Number(r[ART_COL_BLASTO]) || 0;
    });
    var mesesArt = Object.keys(artPorMes).sort().slice(-6);

    // ── 2. Leer FET (encabezado en fila 2, datos desde fila 3) ──
    var shFet  = findSheet(ssLab, 'FET');
    var fetRaw = shFet ? shFet.getDataRange().getValues() : [];
    var fetData = fetRaw.slice(2).filter(function(r) {
      return r.some(function(c) { return String(c).trim() !== ''; });
    });
    var totalFet      = fetData.length;
    var fetSurvividos = fetData.filter(function(r) {
      return String(r[FET_COL_SURVIVED]).trim().toLowerCase() === 'si';
    }).length;
    var fetPregnancy  = fetData.filter(function(r) {
      return String(r[FET_COL_PREG]).trim().toLowerCase() === 'si';
    }).length;

    // Tendencia FET por mes
    var fetPorMes = {};
    fetData.forEach(function(r) {
      var mes = mesLabel(r[FET_COL_FECHA]);
      if (!mes) return;
      if (!fetPorMes[mes]) fetPorMes[mes] = { total:0, survived:0 };
      fetPorMes[mes].total++;
      if (String(r[FET_COL_SURVIVED]).trim().toLowerCase() === 'si') fetPorMes[mes].survived++;
    });

    // ── 3. Inventario Crío (encabezado fila 1, datos desde fila 2) ──
    var shCrio  = findSheet(ssLab, 'Inventario Crío');
    var crioRaw = shCrio ? shCrio.getDataRange().getValues() : [];
    var crioData = crioRaw.slice(1).filter(function(r) {
      return String(r[CRIO_COL_NOMBRE]).trim() !== '';
    });
    var totalOvCrio  = crioData.reduce(function(s,r){ return s+(Number(r[CRIO_COL_OOV])||0); },0);
    var totalEmbCrio = crioData.reduce(function(s,r){ return s+(Number(r[CRIO_COL_EMB])||0); },0);

    // ── 4. Insumos (encabezado fila 2, datos desde fila 3) ──
    var shIns  = findSheet(ssLab, 'Insumos');
    var insRaw = shIns ? shIns.getDataRange().getValues() : [];
    var insData = insRaw.slice(2).filter(function(r) {
      return String(r[INS_COL_INSUMO]).trim() !== '';
    });

    // ── Construir respuesta ──
    var fetPct = pct(fetSurvividos, totalFet);
    var fecPct = (ART_COL_2PN > -1) ? pct(total2PN, totalOocitos) : null;
    var blaPct = (ART_COL_BLASTO > -1) ? pct(totalBlasto, total2PN) : null;

    return {
      view:   'lab-resumen',
      fuente: 'lab-resumen',
      kpis: {
        fecundacion:      fecPct,
        blastocistos:     blaPct,
        fetSupervivencia: fetPct,
        icsiDano:         null,           // configurar ART_COL_ICSI_DANO
        totalCiclos:      totalCiclos,
        totalFet:         totalFet,
        fetPregnancy:     fetPregnancy,
        totalOvCrio:      totalOvCrio,
        totalEmbCrio:     totalEmbCrio,
        mesPeriodo:       mesesArt.length ? mesesArt[mesesArt.length-1] : ''
      },
      tendencia: {
        meses:        mesesArt,
        ciclos:       mesesArt.map(function(m){ return (artPorMes[m]||{}).ciclos||0; }),
        fecundacion:  mesesArt.map(function(m){
          var d = artPorMes[m]||{}; return ART_COL_2PN>-1 ? pct(d.pn2||0, d.oocitos||0) : null;
        }),
        fetSupervivencia: mesesArt.map(function(m){
          var d = fetPorMes[m]||{}; return pct(d.survived||0, d.total||0);
        })
      },
      tanques: crioData.slice(0,10).map(function(r) {
        return {
          nombre: String(r[CRIO_COL_NOMBRE]).split(' ').slice(0,2).join(' '),
          oocitos: Number(r[CRIO_COL_OOV])||0,
          embriones: Number(r[CRIO_COL_EMB])||0
        };
      }),
      insumos: insData.map(function(r) {
        return {
          item:        String(r[INS_COL_INSUMO]),
          proveedor:   String(r[INS_COL_PROV] || ''),
          fecha:       fmtFecha(r[INS_COL_FECHA]),
          costo:       Number(r[INS_COL_COSTO]) || 0,
          estado:      'ok'
        };
      })
    };
  } catch(ex) {
    return { view: 'lab-resumen', fuente: 'lab-resumen', error: ex.message,
             kpis: {}, tendencia: { meses:[], ciclos:[], fecundacion:[], fetSupervivencia:[] },
             tanques: [], insumos: [] };
  }
}

/* ══ DASHBOARD QUIROFANO ════════════════════════════════════════
   Placeholder — ampliar cuando se cree el spreadsheet de Qx
   ══════════════════════════════════════════════════════════════ */
function readQxResumen(fechaInicio, fechaFin) {
  try {
    var ssQx = SpreadsheetApp.openById(QX_SS_ID);
    // Intentar leer hoja "Insumos Qx" para conteos básicos
    var insSheet = findSheet(ssQx, 'Insumos Qx');
    var totalInsumos = 0, totalCosto = 0;
    if (insSheet) {
      var allRows = insSheet.getDataRange().getValues();
      var hdrs = allRows[0];
      var colCosto = hdrs.indexOf('Costo');
      allRows.slice(1).forEach(function(r) {
        if (r[0]) {
          totalInsumos++;
          if (colCosto >= 0 && r[colCosto]) totalCosto += parseFloat(r[colCosto]) || 0;
        }
      });
    }
    return {
      view: 'qx-resumen', fuente: 'qx-resumen',
      kpis: [
        { label: 'Insumos registrados', value: totalInsumos, format: 'number', icon: 'package' },
        { label: 'Costo total insumos', value: totalCosto,   format: 'currency', icon: 'dollar-sign' }
      ],
      rows: [], headers: []
    };
  } catch(ex) {
    return { view: 'qx-resumen', fuente: 'qx-resumen', kpis: [], rows: [], headers: [], error: ex.message };
  }
}

/* ══ REPORTE EJECUTIVO — Agrega KPIs de todas las secciones ═══ */
function readRepEjecutivo(ss, fechaInicio, fechaFin) {
  var result = {
    view: 'rep-ejecutivo', fuente: 'Rep Ejecutivo',
    periodo: (fechaInicio || '') + ' — ' + (fechaFin || ''),
    sections: { financiero: {}, clinico: {}, operaciones: {} },
    rows: [], headers: []
  };

  // ── FINANCIERO: leer Mensual_Todos filtrado por período ──────
  try {
    var mensualSheet = findSheet(ss, 'Mensual_Todos');
    if (mensualSheet) {
      var mRows = mensualSheet.getDataRange().getValues();
      var mHdrs = mRows[0];
      var colPeriodo  = mHdrs.indexOf('Periodo');
      var colIngresos = mHdrs.indexOf('Ingresos');
      var colGastos   = mHdrs.indexOf('Gastos');
      var totIng = 0, totGas = 0, found = false;
      mRows.slice(1).forEach(function(r) {
        var p = colPeriodo >= 0 ? String(r[colPeriodo]) : '';
        if (fechaInicio && p < fechaInicio.slice(0,7)) return;
        if (fechaFin   && p > fechaFin.slice(0,7))   return;
        if (colIngresos >= 0) totIng += parseFloat(r[colIngresos]) || 0;
        if (colGastos   >= 0) totGas += parseFloat(r[colGastos])   || 0;
        found = true;
      });
      if (found || totIng || totGas) {
        result.sections.financiero.ingresos = totIng;
        result.sections.financiero.gastos   = totGas;
      }
    }
  } catch(e) {}

  // ── CLÍNICO: leer Lab SS ──────────────────────────────────────
  try {
    var ssLab = SpreadsheetApp.openById(LAB_SS_ID);

    var artSheet = findSheet(ssLab, 'ART Lab');
    if (artSheet) {
      var artRows = artSheet.getDataRange().getValues().slice(2);
      result.sections.clinico.ciclosART = artRows.filter(function(r){ return r[0]; }).length;
    }

    var fetSheet = findSheet(ssLab, 'FET');
    if (fetSheet) {
      var fetRows = fetSheet.getDataRange().getValues().slice(2);
      result.sections.clinico.fetRealizados = fetRows.filter(function(r){ return r[0]; }).length;
    }

    // Banco Crío
    var crioSheet = findSheet(ssLab, 'Inventario Crío') || findSheet(ssLab, 'Inventario Crio');
    if (crioSheet) {
      var crioRows = crioSheet.getDataRange().getValues().slice(1);
      result.sections.operaciones.bancoCrio = crioRows.filter(function(r){ return r[0]; }).length;
    }

    // Insumos activos (Lab)
    var insSheet = findSheet(ssLab, 'Insumos');
    if (insSheet) {
      var insRows = insSheet.getDataRange().getValues().slice(1);
      result.sections.operaciones.insumosActivos = insRows.filter(function(r){ return r[0]; }).length;
    }
  } catch(e) {}

  // ── OPERACIONES: alertas de inventario desde hoja Alertas ────
  try {
    var alertSheet = findSheet(ss, 'Alertas');
    if (alertSheet) {
      var alertRows = alertSheet.getDataRange().getValues().slice(1);
      result.sections.operaciones.alertasInventario = alertRows.filter(function(r){ return r[0]; }).length;
    }
  } catch(e) {}

  return result;
}

/* ══ DASHBOARD MEDICAMENTOS ════════════════════════════════════ */
function readMedDashboard(ss, fechaInicio, fechaFin) {
  var medId = CAPTURA_SHEETS['Medicamentos'] || CAPTURA_SHEET_ID_DEFAULT;
  var ssMed = SpreadsheetApp.openById(medId);

  function readSheet(nombre) {
    var sh = findSheet(ssMed, nombre);
    if (!sh || sh.getLastRow() < 2) return { headers: [], rows: [] };
    var vals = sh.getDataRange().getValues();
    var hdrs = vals[0].map(function(h){ return String(h).trim(); });
    var rows = vals.slice(1).filter(function(r){
      return r.some(function(c){ return String(c).trim() !== ''; });
    }).map(function(r){
      var obj = {};
      hdrs.forEach(function(h, i){ obj[h] = r[i]; });
      return obj;
    });
    // Filtrar por fecha si existe columna Fecha
    var tieneFecha = hdrs.indexOf('Fecha') !== -1;
    if (tieneFecha && fechaInicio && fechaFin) {
      rows = rows.filter(function(r){
        var raw = r['Fecha'];
        var f = (raw instanceof Date)
          ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(raw || '').slice(0, 10);
        return f >= fechaInicio && f <= fechaFin;
      });
    }
    return { headers: hdrs, rows: rows };
  }

  var compraData = readSheet('Ent. Med');
  var estimData  = readSheet('Estimulación');
  var compras    = compraData.rows;
  var estims     = estimData.rows;

  // ── Agregaciones compras ──────────────────────────────────────
  var comprasPorMed = {}, gastosPorMed = {}, comprasPorMes = {};
  compras.forEach(function(r) {
    var med = String(r['Medicamento'] || '').trim();
    var cant = Number(r['Cantidad']) || 0;
    var total = Number(r['Total']) || (cant * (Number(r['Precio_Unitario']) || 0));
    var mes = String(r['Fecha'] || '').slice(0, 7);
    if (med) {
      comprasPorMed[med] = (comprasPorMed[med] || 0) + cant;
      gastosPorMed[med]  = (gastosPorMed[med]  || 0) + total;
    }
    if (mes) comprasPorMes[mes] = (comprasPorMes[mes] || 0) + cant;
  });

  // ── Columnas de medicamentos: todo lo que viene DESPUÉS de "Cancelado" ──
  // Al agregar nuevas columnas en Sheets después de Cancelado se incluyen automáticamente
  var estimHeaders = estimData.headers || [];
  var canceladoIdx = -1;
  for (var ci = 0; ci < estimHeaders.length; ci++) {
    if (estimHeaders[ci].trim().toLowerCase() === 'cancelado') { canceladoIdx = ci; break; }
  }
  var MED_COLS = canceladoIdx >= 0
    ? estimHeaders.slice(canceladoIdx + 1).filter(function(h){ return h.trim() !== ''; })
    : [];
  var usosPorMed = {}, usosPorMes = {}, pacientesSet = {};
  estims.forEach(function(r) {
    var pac = String(r['Paciente'] || '').trim();
    var mes = String(r['Fecha']    || '').slice(0, 7);
    if (pac) pacientesSet[pac] = 1;
    var totalFila = 0;
    MED_COLS.forEach(function(col) {
      var cant = Number(r[col]) || 0;
      if (cant > 0) {
        usosPorMed[col] = (usosPorMed[col] || 0) + cant;
        totalFila += cant;
      }
    });
    if (mes && totalFila > 0) usosPorMes[mes] = (usosPorMes[mes] || 0) + totalFila;
  });

  // ── Top 8 ─────────────────────────────────────────────────────
  function top8(obj) {
    return Object.keys(obj).map(function(k){ return { label: k, value: obj[k] }; })
      .sort(function(a,b){ return b.value - a.value; }).slice(0, 8);
  }

  var topCompras = top8(comprasPorMed);
  var topUsos    = top8(usosPorMed);

  // ── Evolución mensual ─────────────────────────────────────────
  var mesesSet = {};
  Object.keys(comprasPorMes).forEach(function(m){ mesesSet[m]=1; });
  Object.keys(usosPorMes).forEach(function(m){ mesesSet[m]=1; });
  var meses = Object.keys(mesesSet).sort();

  // ── KPIs ──────────────────────────────────────────────────────
  var totalCompras = compras.reduce(function(s,r){ return s+(Number(r['Cantidad'])||0); }, 0);
  var totalUsos    = Object.values ? Object.values(usosPorMed).reduce(function(s,v){ return s+v; }, 0)
                    : Object.keys(usosPorMed).reduce(function(s,k){ return s+usosPorMed[k]; }, 0);
  var gastoTotal   = estims.reduce(function(s,r){ return s+(Number(r['Costo Meds'])||0); }, 0)
                    + compras.reduce(function(s,r){
    return s + (Number(r['Total']) || (Number(r['Cantidad'])||0)*(Number(r['Precio_Unitario'])||0));
  }, 0);

  return {
    view:   'med-resumen',
    periodo: fechaInicio.slice(0,7) + ' → ' + fechaFin.slice(0,7),
    kpis: {
      totalCompras:       totalCompras,
      totalUsos:          totalUsos,
      gastoTotal:         gastoTotal,
      medsDistintos:      Object.keys(comprasPorMed).length,
      pacientesAtendidos: Object.keys(pacientesSet).length
    },
    topCompras: topCompras,
    topUsos:    topUsos,
    evolucion: {
      meses:   meses,
      compras: meses.map(function(m){ return comprasPorMes[m] || 0; }),
      usos:    meses.map(function(m){ return usosPorMes[m]    || 0; })
    }
  };
}

/* ── Datos financieros completos (filtrados por rango de fechas) ─── */
function readMensualData(ss, fechaInicio, fechaFin, viewId) {
  var label = fechaInicio.slice(0,7) + ' → ' + fechaFin.slice(0,7);
  return {
    view:          viewId,
    periodo:       label,          // string descriptivo para subtítulos
    fechaInicio:   fechaInicio,
    fechaFin:      fechaFin,
    todos:         readMensual(ss, 'Mensual_Todos',         fechaInicio, fechaFin),
    local:         readMensual(ss, 'Mensual_Local',         fechaInicio, fechaFin),
    internacional: readMensual(ss, 'Mensual_Internacional', fechaInicio, fechaFin),
    servicios:     readServicios(ss),
    funnel:        readFunnel(ss),
    alertas:       readAlertas(ss),
    donut:         readDonut(ss),
    cashflow:      readCashFlow(ss, fechaInicio, fechaFin),
    costos:        readCostos(ss),
    paisesOrigen:  readPaisesOrigen(ss, fechaInicio, fechaFin),
    updated:       new Date().toISOString()
  };
}

/* ── Datos de hoja de captura — devuelve TODAS las filas (sin filtro de fechas)
   Las hojas de captura (Pacientes, Productos, Ent. Med, Estimulacion, etc.)
   muestran su contenido completo; el filtro de fechas aplica solo a datos financieros.
   ──────────────────────────────────────────────────────────────── */
function findSheet(ssCap, nombreHoja) {
  // 1. Nombre exacto
  var h = ssCap.getSheetByName(nombreHoja);
  if (h) return h;
  // 2. Búsqueda insensible a tildes y mayúsculas
  var normalize = function(s) {
    return s.trim().toLowerCase()
      .replace(/[áàäã]/g,'a').replace(/[éèë]/g,'e')
      .replace(/[íìï]/g,'i').replace(/[óòö]/g,'o')
      .replace(/[úùü]/g,'u').replace(/ñ/g,'n');
  };
  var target = normalize(nombreHoja);
  var sheets = ssCap.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalize(sheets[i].getName()) === target) return sheets[i];
  }
  return null;
}

function getCapturaId(nombreHoja) {
  if (CAPTURA_SHEETS[nombreHoja]) return CAPTURA_SHEETS[nombreHoja];
  // Búsqueda tolerante a tildes
  var normalize = function(s) {
    return s.trim().toLowerCase()
      .replace(/[áàäã]/g,'a').replace(/[éèë]/g,'e')
      .replace(/[íìï]/g,'i').replace(/[óòö]/g,'o')
      .replace(/[úùü]/g,'u').replace(/ñ/g,'n');
  };
  var target = normalize(nombreHoja);
  var keys = Object.keys(CAPTURA_SHEETS);
  for (var i = 0; i < keys.length; i++) {
    if (normalize(keys[i]) === target) return CAPTURA_SHEETS[keys[i]];
  }
  return CAPTURA_SHEET_ID_DEFAULT;
}

function readCapturaData(ss, nombreHoja, viewId, fechaInicio, fechaFin) {
  var capturaId = getCapturaId(nombreHoja);
  var ssCap = SpreadsheetApp.openById(capturaId);
  var hoja  = findSheet(ssCap, nombreHoja);
  if (!hoja) {
    return { view: viewId, headers: [], rows: [],
             error: 'Hoja "' + nombreHoja + '" no encontrada.' };
  }
  var allRows = hoja.getDataRange().getValues();
  if (allRows.length < 1) return { view: viewId, headers: [], rows: [] };

  // ── Detección inteligente de fila de encabezado ────────────────
  // Algunas hojas tienen fila 1 vacía (FET, Insumos) o encabezados
  // divididos entre fila 1 y fila 2 (ART Lab). Se detecta automáticamente.
  function countFilled(row) {
    return row.filter(function(c) { return String(c).trim() !== ''; }).length;
  }
  var r0 = countFilled(allRows[0]);
  var r1 = allRows.length > 1 ? countFilled(allRows[1]) : 0;
  var headerRow, dataStart;

  if (r0 === 0) {
    // Fila 1 vacía → encabezado en fila 2 (FET, Insumos)
    headerRow  = allRows[1] || [];
    dataStart  = 2;
  } else if (r0 > 0 && r1 > 0) {
    // Ambas filas tienen datos → verificar si son encabezados complementarios
    // (sin solapamiento de celdas llenas, patrón ART Lab)
    var complementario = allRows[0].every(function(v, i) {
      var v0 = String(v).trim();
      var v1 = String((allRows[1][i] !== undefined ? allRows[1][i] : '')).trim();
      return !(v0 && v1); // No hay posición donde ambas filas tengan valor
    });
    if (complementario) {
      // Fusionar fila 1 y fila 2 como encabezado único (ART Lab)
      headerRow = allRows[0].map(function(v, i) {
        return String(v).trim() || String(allRows[1][i] !== undefined ? allRows[1][i] : '').trim();
      });
      dataStart = 2;
    } else {
      // Fila 1 tiene encabezados reales; fila 2 es la primera fila de datos
      headerRow = allRows[0];
      dataStart = 1;
    }
  } else {
    // Caso normal: fila 1 = encabezados
    headerRow = allRows[0];
    dataStart = 1;
  }

  // Detectar columna Periodo oculta en col A (se excluye de la vista)
  var tienePeriodo = String(headerRow[0]).trim().toLowerCase() === 'periodo';
  var colStart = tienePeriodo ? 1 : 0;
  var headers = headerRow.slice(colStart)
    .map(function(h) { return String(h).trim(); })
    .filter(function(h) { return h !== ''; });

  // Incluir todas las filas no vacías (sin filtro de fechas)
  var dataRowsWithNum = allRows.slice(dataStart)
    .map(function(r, i) { return { data: r, rowNum: i + dataStart + 1 }; })
    .filter(function(item) {
      return item.data.some(function(c) { return String(c).trim() !== ''; });
    });

  var rows = dataRowsWithNum.map(function(item) {
    var r = item.data;
    var obj = { _rowNum: item.rowNum, _periodo: tienePeriodo ? String(r[0]) : '' };
    headers.forEach(function(h, i) {
      obj[h] = r[colStart + i];
      obj[h.toLowerCase()] = r[colStart + i];
    });
    return obj;
  });

  return { view: viewId, fuente: nombreHoja, headers: headers, rows: rows,
           updated: new Date().toISOString() };
}

/* ══════════════════════════════════════════════════════════════
   LECTORES INDIVIDUALES
   ══════════════════════════════════════════════════════════════ */
function readPeriodos(ss) {
  var rows = ss.getSheetByName('Periodos').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { id: String(r[0]), label: String(r[1]), orden: Number(r[2]) };
  });
}

/* Columnas esperadas (con la nueva col Fecha en B):
   A=Periodo | B=Fecha(YYYY-MM-DD) | C=Mes | D=Ingresos | E=Gastos | F=Ciclos | G=CAC | H=Margen */
function readMensual(ss, sheetName, fechaInicio, fechaFin) {
  var rows = ss.getSheetByName(sheetName).getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[1]).trim(); // col B = Fecha
    return f >= fechaInicio && f <= fechaFin;
  });
  // Ordenar por fecha ascendente
  data.sort(function(a, b) { return String(a[1]) < String(b[1]) ? -1 : 1; });
  return {
    meses:    data.map(function(r) { return String(r[2]); }),  // col C
    ingresos: data.map(function(r) { return Number(r[3]); }),  // col D
    gastos:   data.map(function(r) { return Number(r[4]); }),  // col E
    ciclos:   data.map(function(r) { return Number(r[5]); }),  // col F
    cac:      data.map(function(r) { return Number(r[6]); }),  // col G
    margen:   data.map(function(r) { return Number(r[7]); })   // col H
  };
}

/* Columnas CashFlow: A=Periodo | B=Fecha | C=Mes | D=Flujo_MXN */
function readCashFlow(ss, fechaInicio, fechaFin) {
  var rows = ss.getSheetByName('CashFlow').getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[1]).trim();
    return f >= fechaInicio && f <= fechaFin;
  });
  data.sort(function(a, b) { return String(a[1]) < String(b[1]) ? -1 : 1; });
  return {
    meses: data.map(function(r) { return String(r[2]); }),
    flujo: data.map(function(r) { return Number(r[3]); })
  };
}

function readServicios(ss) {
  var rows = ss.getSheetByName('Servicios').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { name: String(r[0]), color: String(r[1]), ingresos: String(r[2]),
             margen: Number(r[3]), meta: Number(r[4]) };
  });
}

function readFunnel(ss) {
  var rows = ss.getSheetByName('Funnel').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { label: String(r[0]), val: Number(r[1]), pct: Number(r[2]), color: String(r[3]) };
  });
}

function readAlertas(ss) {
  var rows = ss.getSheetByName('Alertas').getDataRange().getValues();
  return rows.slice(1).map(function(r) {
    return { type: String(r[0]), icon: String(r[1]), title: String(r[2]), desc: String(r[3]) };
  });
}

function readDonut(ss) {
  var rows = ss.getSheetByName('DonutServicios').getDataRange().getValues();
  var data = rows.slice(1);
  return {
    labels: data.map(function(r) { return String(r[0]); }),
    data:   data.map(function(r) { return Number(r[1]); }),
    colors: data.map(function(r) { return String(r[2]); })
  };
}

/* PaisesOrigen: A=Periodo | B=Fecha | C=Pais | D=Porcentaje | E=Color */
function readPaisesOrigen(ss, fechaInicio, fechaFin) {
  var hoja = ss.getSheetByName('PaisesOrigen');
  if (!hoja) return { labels: [], data: [], colors: [] };
  var rows = hoja.getDataRange().getValues();
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[1]).trim();
    return f >= fechaInicio && f <= fechaFin;
  });
  return {
    labels: data.map(function(r) { return String(r[2]); }),
    data:   data.map(function(r) { return Number(r[3]); }),
    colors: data.map(function(r) { return String(r[4]); })
  };
}

/* ══════════════════════════════════════════════════════════════
   INSERT ROW — agrega una fila al final de la hoja indicada
   Params: sheet, periodo, + una clave por columna (según cabecera)
   ══════════════════════════════════════════════════════════════ */
/* Detecta la fila de encabezados real de una hoja (maneja row1-vacía y headers combinados row1+row2) */
function getSheetHeaders(hoja) {
  var numCols = hoja.getLastColumn();
  var numRows = Math.min(hoja.getLastRow(), 3);
  if (numRows === 0 || numCols === 0) return { headers: [], dataStart: 1 };
  var allRows = hoja.getRange(1, 1, numRows, numCols).getValues();
  function countFilled(row) {
    return row.filter(function(c) { return String(c).trim() !== ''; }).length;
  }
  var r0 = countFilled(allRows[0]);
  var r1 = allRows.length > 1 ? countFilled(allRows[1]) : 0;
  var headers, dataStart;
  if (r0 === 0) {
    headers = allRows[1] || []; dataStart = 3;
  } else if (r0 > 0 && r1 > 0) {
    var complementario = allRows[0].every(function(v, i) {
      var v0 = String(v).trim();
      var v1 = String((allRows[1][i] !== undefined ? allRows[1][i] : '')).trim();
      return !(v0 && v1);
    });
    if (complementario) {
      headers = allRows[0].map(function(v, i) {
        return String(v).trim() || String(allRows[1][i] !== undefined ? allRows[1][i] : '').trim();
      });
      dataStart = 3;
    } else { headers = allRows[0]; dataStart = 2; }
  } else { headers = allRows[0]; dataStart = 2; }
  return { headers: headers.map(function(h){ return String(h).trim(); }), dataStart: dataStart };
}

function insertRow(ss, e) {
  var sheetName = (e && e.parameter.sheet) || '';
  if (SHEET_ALIASES[sheetName]) sheetName = SHEET_ALIASES[sheetName];
  var capturaId = getCapturaId(sheetName);
  var ssCap = SpreadsheetApp.openById(capturaId);
  var hoja = findSheet(ssCap, sheetName);
  if (!hoja) return { error: 'Hoja "' + sheetName + '" no encontrada.' };

  var hdrInfo = getSheetHeaders(hoja);
  var headers = hdrInfo.headers;
  var row = headers.map(function(h, i) {
    if (i === 0 && h === 'Periodo') return (e && e.parameter.periodo) || '';
    return (e && e.parameter[h] !== undefined) ? e.parameter[h] : '';
  });

  hoja.appendRow(row);
  invalidateViewCache(sheetName);
  return { success: true };
}

function readCostos(ss) {
  var rows = ss.getSheetByName('DistribucionCostos').getDataRange().getValues();
  var data = rows.slice(1);
  return {
    labels: data.map(function(r) { return String(r[0]); }),
    data:   data.map(function(r) { return Number(r[1]); }),
    colors: data.map(function(r) { return String(r[2]); })
  };
}

/* ══════════════════════════════════════════════════════════════
   setupSheets — Ejecutar para crear/migrar todas las hojas
   ══════════════════════════════════════════════════════════════ */
function setupSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Hoja: Menu ──────────────────────────────────────────────
  // Columnas: ID | Padre | Label | Seccion | Icono | Orden | Tipo | Fuente | Activo
  crearHoja(ss, 'Menu', [
    ['ID',            'Padre',     'Label',            'Seccion',  'Icono',            'Orden', 'Tipo',  'Fuente',   'Activo'],
    // ── PANEL ──
    ['resumen',       '',          'Resumen General',  'PANEL',    'layout-dashboard', 1,       'vista', 'mensual',  true],
    ['finanzas',      '',          'Finanzas',         'PANEL',    'landmark',         2,       'grupo', '',         true],
    ['ingresos',      'finanzas',  'Ingresos',         '',         'trending-up',      1,       'vista', 'mensual',  true],
    ['gastos',        'finanzas',  'Gastos Operativos','',         'receipt',          2,       'vista', 'mensual',  true],
    ['costos',        'finanzas',  'Costos',           '',         'calculator',       3,       'vista', 'mensual',  true],
    ['pacientes',     '',          'Pacientes',        'PANEL',    'users',            3,       'vista', 'mensual',  true],
    // ── ANÁLISIS ──
    ['analisis',      '',          'Análisis',         'ANÁLISIS', 'bar-chart-2',      4,       'grupo', '',         true],
    ['servicios',     'analisis',  'Servicios',        '',         'flask-conical',    1,       'vista', 'mensual',  true],
    ['turismo',       'analisis',  'Turismo Médico',   '',         'plane',            2,       'vista', 'mensual',  true],
    ['rentabilidad',  'analisis',  'Rentabilidad',     '',         'percent',          3,       'vista', 'mensual',  true],
    // ── CAPTURA ──
    ['medicamentos',  '',          'Medicamentos',     'PANEL',    'pill',             4,       'grupo', '',              true],
    ['med-resumen',   'medicamentos','Resumen Med.',   '',         'bar-chart-2',      1,       'vista', 'med-dashboard', true],
    ['captura',       '',          'Captura',          'CAPTURA',  'database',         5,       'grupo', '',         true],
    ['inventarios',   'captura',   'Inventarios',      '',         'package',          1,       'vista', 'Inventarios',  true],
    ['laboratorios',  'captura',   'Laboratorios',     '',         'microscope',       2,       'vista', 'Laboratorios', true],
    // ── CONFIG ──
    ['ajustes',       '',          'Ajustes',          'CONFIG',   'settings',         6,       'vista', '',         true],
  ]);

  // ── Hoja: Periodos ──────────────────────────────────────────
  crearHoja(ss, 'Periodos', [
    ['ID',         'Label',      'Orden'],
    ['2026-Q2',    'Q2 2026',    1],
    ['2026-Q1',    'Q1 2026',    2],
    ['2025-Anual', '2025 Anual', 3],
  ]);

  Logger.log('✅ setupSheets completado');
}

/* ══════════════════════════════════════════════════════════════
   GESTIÓN DE ROLES — Lee todos los roles con sus permisos
   ══════════════════════════════════════════════════════════════ */
function readGestionRoles(ss) {
  var shR = ss.getSheetByName('Roles');
  var roles = [];
  if (shR && shR.getLastRow() > 1) {
    var data = shR.getDataRange().getValues();
    var h    = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
    var rI   = h.indexOf('rol');
    var bI   = h.indexOf('vistas_bloqueadas');
    var lI   = h.indexOf('solo_lectura');
    var dI   = h.indexOf('descripcion');
    for (var i = 1; i < data.length; i++) {
      var rolName = rI > -1 ? String(data[i][rI]).trim() : '';
      if (!rolName) continue;
      var bloq = bI > -1 ? String(data[i][bI]).trim() : '';
      roles.push({
        rol:              rolName,
        vistasBloqueadas: bloq ? bloq.split(',').map(function(v){ return v.trim(); }).filter(Boolean) : [],
        soloLectura:      lI > -1 ? !!data[i][lI] : false,
        descripcion:      dI > -1 ? String(data[i][dI]).trim() : ''
      });
    }
  }
  return { view: 'gestion-roles', fuente: 'Roles', roles: roles };
}

function crearHoja(ss, nombre, datos) {
  var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  h.clearContents();
  h.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
  var header = h.getRange(1, 1, 1, datos[0].length);
  header.setFontWeight('bold');
  header.setBackground('#fce8f0');
  h.setFrozenRows(1);
  h.autoResizeColumns(1, datos[0].length);
}
