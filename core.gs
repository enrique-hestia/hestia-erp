/* ==============================================================
   core.gs — Núcleo del ERP
   --------------------------------------------------------------
   Auth, doGet, doPost, menú, router, readViewData
   Proyecto Google Apps Script — Hestia Fertility ERP
   Todas las constantes vienen de config.gs (mismo proyecto)
   ============================================================== */

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
// Valida credenciales y devuelve token + permisos. Compartida por doGet y doPost.
function handleLogin(email, password) {
  email = email || '';
  if (!email) return { error: 'Email requerido.' };
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var shU = ss.getSheetByName('Usuarios');
  if (!shU) return { error: 'Módulo de usuarios no configurado.' };
  var user = getUserRow(ss, email);
  if (!user)        return { error: 'Usuario no encontrado.' };
  if (!user.activo) return { error: 'Usuario inactivo. Contacta al administrador.' };
  if (user.password && user.password !== password)
                    return { error: 'Contraseña incorrecta.' };
  var rolCfg = getRolConfig(ss, user.rol);
  return {
    success: true, token: generateToken(user.email),
    email: user.email, nombre: user.nombre, rol: user.rol,
    vistasBloqueadas: rolCfg.vistasBloqueadas,
    soloLectura:      rolCfg.soloLectura
  };
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

    // login via GET (POST body se pierde en redirect de GAS)
    if (action === 'login') {
      var loginEmail = (e && e.parameter.email) || '';
      var loginPass  = (e && e.parameter.password) || '';
      return jsonResponse(handleLogin(loginEmail, loginPass));
    }

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
    var sucursal    = (e && e.parameter.sucursal)    || 'Todas';
    var viewType    = (e && e.parameter.viewType)    || '';
    var plMonth     = (e && e.parameter.plMonth)     || '';
    var plYear      = (e && e.parameter.plYear)      || '';
    var plPrevYear  = (e && e.parameter.plPrevYear)  || '';

    // ── LOGIN: valida credenciales y devuelve token + permisos ──
    // Se mantiene por GET por compatibilidad, pero el frontend ahora usa POST
    // (la contraseña no debe viajar en la URL). Ver handleLogin() abajo.
    if (action === 'login') {
      return jsonResponse(handleLogin(
        (e && e.parameter.email)    || '',
        (e && e.parameter.password) || ''
      ));
    }

    // ── BYPASS: lectura de egresos con API key dedicada (integraciones externas, p.ej. Claude) ──
    // Clave propia e independiente de AUTH_SECRET para evitar depender de cuál definición
    // de AUTH_SECRET esté activa en el proyecto (hardcoded vs. Script Properties).
    if (action === 'egresos_raw') {
      var ECRAW_KEY = 'hestia-ecraw-9f2a71';
      var apiKey = (e && e.parameter.key) || '';
      if (!apiKey || apiKey !== ECRAW_KEY) {
        return jsonResponse({ error: 'API key inválida.', code: 401 });
      }
      var anioRaw = parseInt((e && e.parameter.anio) || new Date().getFullYear());
      var fpRaw   = (e && e.parameter.fp) || '';
      var egData  = readEgresosData(anioRaw);
      var rows    = egData.rows || [];
      if (fpRaw) rows = rows.filter(function(r){ return r.formaPago === fpRaw; });
      return jsonResponse({ ok: true, egresos: rows, anio: anioRaw, fp: fpRaw, total: rows.length });
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
      if (!currentUser || String(currentUser.rol||'').toLowerCase() !== 'admin')
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

    // ── SAVEUSER: se maneja SOLO por POST (finance.gs doPost) para no
    //    pasar la contraseña por la URL ni borrarla al editar. ──

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

    if (action === 'banks') {
      var banksFi       = (e && e.parameter.fechaInicio) || '';
      var banksFf       = (e && e.parameter.fechaFin)    || '';
      var banksCache    = CacheService.getScriptCache();
      var banksCacheKey = 'erp_banks_v3_' + banksFi + '_' + banksFf;
      var banksCached   = banksCache.get(banksCacheKey);
      if (banksCached) {
        return ContentService.createTextOutput(banksCached).setMimeType(ContentService.MimeType.JSON);
      }
      var banksResult = readBanksData(banksFi, banksFf);
      var banksJson   = JSON.stringify(banksResult);
      try { banksCache.put(banksCacheKey, banksJson, 30); } catch(e) {}
      return ContentService.createTextOutput(banksJson).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'menu') {
      var menuCache    = CacheService.getScriptCache();
      var menuCacheKey = 'erp_menu_v3_' + fechaInicio + '_' + fechaFin;
      var menuCached   = menuCache.get(menuCacheKey);
      if (menuCached) {
        return ContentService.createTextOutput(menuCached).setMimeType(ContentService.MimeType.JSON);
      }
      var menuResult = {
        menu:        readMenu(ss),
        sucursales:  readSucursales(ss),
        fechaInicio: fechaInicio,
        fechaFin:    fechaFin,
        version:     API_VERSION
      };
      var menuJson = JSON.stringify(menuResult);
      try { menuCache.put(menuCacheKey, menuJson, 1800); } catch(e) {}
      return ContentService.createTextOutput(menuJson).setMimeType(ContentService.MimeType.JSON);
    }

    // Caja Chica: dashboard con saldo y resumen de gasto
    if (action === 'cajachica') {
      return jsonResponse(readCajaChicaData());
    }

    // Cuentas por Pagar: reporte de vencimientos
    if (action === 'cxp') {
      return jsonResponse(readCxPData());
    }

    if (action === 'bdcxp') {
      return jsonResponse(readBDCxP());
    }

    // Gastos fijos (recurrentes): catálogo, propuestas del mes, proyección
    if (action === 'gfAll') {
      var _per = e.parameter.periodo || '';
      return jsonResponse({ok:true, propuestas:readGastosFijosPropuestas(_per), catalogo:readGastosFijos()});
    }
    if (action === 'gastosFijos') {
      return jsonResponse(readGastosFijos());
    }
    if (action === 'gfPropuestas') {
      return jsonResponse(readGastosFijosPropuestas(e.parameter.periodo || ''));
    }
    if (action === 'gfProyeccion') {
      return jsonResponse(readProyeccionGastosFijos({periodo:e.parameter.periodo||'', meses:e.parameter.meses||3}));
    }

    // Tareas programadas (scheduler) — Panel de Control
    if (action === 'scheduledTasks') {
      return jsonResponse(readScheduledTasks());
    }

    // Recordatorios (agenda personal)
    if (action === 'recordatorios') {
      return jsonResponse(readRecordatorios(e.parameter.usuario || ''));
    }
    if (action === 'recordatoriosPend') {
      return jsonResponse(readRecordatoriosPendientes(e.parameter.usuario || ''));
    }

    // Catálogo de proveedores (fuente única para dropdowns)
    if (action === 'proveedores') {
      return jsonResponse(readProveedores());
    }

    // Presupuesto: proyección del siguiente trimestre + meta + pace
    if (action === 'presupuesto') {
      return jsonResponse(readPresupuesto());
    }

    // Análisis de Egresos: histórico, ranking, recomendaciones
    if (action === 'analisisEgresos') {
      return jsonResponse(readAnalisisEgresos());
    }
    // Análisis de Ingresos: mejores y más vendidos por grupo/tipo
    if (action === 'analisisIngresos') {
      return jsonResponse(readAnalisisIngresos());
    }
    if (action === 'analisisPacientes') {
      return jsonResponse(readAnalisisPacientes());
    }
    if (action === 'analisisServicios') {
      return jsonResponse(readAnalisisServicios());
    }
    if (action === 'analisisSurrogacy') {
      return jsonResponse(readAnalisisSurrogacy());
    }
    if (action === 'analisisRentabilidad') {
      return jsonResponse(readAnalisisRentabilidad());
    }

    if (action === 'ingresos') {
      return jsonResponse(readIngresosData());
    }

    if (action === 'leerXmlFactura') {
      var fid = (e && e.parameter.fileId) || '';
      return jsonResponse(leerXmlFactura(fid));
    }

    if (action === 'reconciliarFacturas') {
      var fIni = (e && e.parameter.fechaInicio) || '';
      var fFin = (e && e.parameter.fechaFin) || '';
      return jsonResponse(reconciliarFacturasXml(fIni, fFin));
    }

    if (action === 'buscarXmlAmplio') {
      var baIni = (e && e.parameter.fechaInicio) || '';
      var baFin = (e && e.parameter.fechaFin) || '';
      var baPad = (e && e.parameter.padMeses) || '3';
      return jsonResponse(buscarXmlAmplio(baIni, baFin, parseInt(baPad, 10)));
    }

    if (action === 'analizarDatosFiscalesPacientes') {
      var dfIni = (e && e.parameter.fechaInicio) || '';
      var dfFin = (e && e.parameter.fechaFin) || '';
      return jsonResponse(analizarDatosFiscalesPacientes(dfIni, dfFin));
    }

    if (action === 'migratePacientesFiscales') {
      return jsonResponse(migratePacientesFiscales());
    }

    if (action === 'analizarDescuentosFiscales') {
      var dfIni2 = (e && e.parameter.fechaInicio) || '';
      var dfFin2 = (e && e.parameter.fechaFin) || '';
      return jsonResponse(analizarDescuentosFiscales(dfIni2, dfFin2));
    }

    if (action === 'productos') {
      return jsonResponse(readProductos());
    }

    if (action === 'catalogoMedicamentos') {
      return jsonResponse(readCatalogoMedicamentos());
    }

    if (action === 'creditosProveedor') {
      return jsonResponse(readCreditosProveedor((e && e.parameter.proveedor) || ''));
    }

    if (action === 'trazaCancelacion') {
      if (typeof readTrazaCancelacion !== 'function')
        return jsonResponse({ok:false, error:'Agrega cxp_creditos.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readTrazaCancelacion((e && e.parameter.cxpId) || ''));
    }

    if (action === 'summary') {
      if (typeof readSummary !== 'function')
        return jsonResponse({ok:false, error:'Agrega summary.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readSummary((e && e.parameter.fechaInicio) || '', (e && e.parameter.fechaFin) || ''));
    }

    if (action === 'summaryConfig') {
      if (typeof readSummaryConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega summary.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readSummaryConfig());
    }

    if (action === 'clasifProveedores') {
      if (typeof readClasificacionProveedores !== 'function')
        return jsonResponse({ok:false, error:'Agrega prov_defaults.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readClasificacionProveedores((e && e.parameter.force) === '1'));
    }

    if (action === 'movimientosInventario') {
      return jsonResponse(readMovimientosInventario({
        sku: (e && e.parameter.sku) || '',
        fechaInicio: (e && e.parameter.fechaInicio) || '',
        fechaFin: (e && e.parameter.fechaFin) || ''
      }));
    }

    if (action === 'combos') {
      return jsonResponse(readCombos());
    }

    if (action === 'ordenesCompra') {
      return jsonResponse(readOrdenesCompra());
    }

    if (action === 'comprobantesEstructura') {
      if (typeof listComprobantesEstructura !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(listComprobantesEstructura((e && e.parameter.fuente) || ''));
    }

    if (action === 'comprobantesMes') {
      if (typeof readComprobantesMes !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readComprobantesMes((e && e.parameter.anio) || '', (e && e.parameter.mes) || '', (e && e.parameter.fuente) || ''));
    }

    if (action === 'plantillasCorreo') {
      if (typeof readPlantillasCorreo !== 'function')
        return jsonResponse({ok:false, error:'Agrega comprobantes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readPlantillasCorreo());
    }

    if (action === 'listas') {
      return jsonResponse(readListas());
    }

    if (action === 'egresos') {
      var anioEg = (e && e.parameter.anio) || '';
      return jsonResponse(readEgresosData(anioEg ? parseInt(anioEg) : undefined));
    }
    if (action === 'egresosSetup') {
      var anioS = parseInt((e && e.parameter.anio) || new Date().getFullYear());
      return jsonResponse(setupEgresosAnio(anioS));
    }

    if (action === 'listaPacientes') {
      return jsonResponse(listaPacientesAll());
    }

    if (action === 'estadoCuenta') {
      var ec = e && e.parameter.paciente ? decodeURIComponent(e.parameter.paciente) : '';
      return jsonResponse(readEstadoCuentaPaciente(ec));
    }

    if (action === 'pacienteLista') {
      var pacNombre = (e && e.parameter.paciente) || '';
      return jsonResponse(readPacienteLista(decodeURIComponent(pacNombre)));
    }

    if (action === 'pacienteFull') {
      if (typeof readPacienteFull !== 'function')
        return jsonResponse({ok:false, error:'Actualiza finance.gs en Apps Script y redespliega.'});
      return jsonResponse(readPacienteFull(decodeURIComponent((e && e.parameter.paciente) || '')));
    }

    if (action === 'getFormatos') {
      return jsonResponse(readFormatos());
    }

    if (action === 'saveFormatos') {
      var fmtData = e.parameter.data || '{}';
      return jsonResponse(saveFormatos(JSON.parse(decodeURIComponent(fmtData))));
    }

    if (action === 'insert') {
      var sheetParamIns = (e && e.parameter.sheet) || '';
      if (sheetParamIns.trim().toLowerCase() === 'caja chica') {
        return jsonResponse(insertCajaChicaRow(e));
      }
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

    // action === 'clearcache' — limpia cache del script (admin)
    if (action === 'clearcache') {
      try { CacheService.getScriptCache().removeAll(['erp_menu_v2_' + fechaInicio + '_' + fechaFin, 'erp_banks_v1']); } catch(e) {}
      return jsonResponse({ ok: true, msg: 'Cache cleared' });
    }

    // action === 'view'  — con caché de 60 s para reducir lecturas a Sheets
    if (action === 'view') {
      var cache    = CacheService.getScriptCache();
      var cacheKey = 'v2_' + view + '_' + fechaInicio + '_' + fechaFin + '_' + sucursal + '_' + viewType + '_' + plMonth + '_' + plYear + '_' + plPrevYear;
      var cached   = cache.get(cacheKey);
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
      var result   = readViewData(ss, view, fechaInicio, fechaFin, sucursal, viewType, plMonth, plYear, plPrevYear);
      var json     = JSON.stringify(result);
      try { cache.put(cacheKey, json, 600); } catch(ignored) {}
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }
    return jsonResponse(readViewData(ss, view, fechaInicio, fechaFin, sucursal, viewType, plMonth, plYear, plPrevYear));

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


/* ══════════════════════════════════════════════════════════════
   LEE LA HOJA Menu
   Columnas: ID | Padre | Label | Seccion | Icono | Orden | Tipo | Fuente | Activo
   ══════════════════════════════════════════════════════════════ */
function readSucursales(ss) {
  var hoja = ss.getSheetByName('Sucursales');
  if (!hoja) return [];
  var rows = hoja.getDataRange().getValues();
  if (rows.length < 2) return [];
  var hdrs = rows[0].map(function(h){ return String(h).trim(); });
  return rows.slice(1)
    .filter(function(r){ return String(r[0]).trim() !== ''; })
    .map(function(r){
      var obj = {};
      hdrs.forEach(function(h, i){ obj[h] = r[i]; });
      return obj;
    })
    .filter(function(s){
      var activo = String(s['Activo'] !== undefined ? s['Activo'] : 'true').trim().toLowerCase();
      return activo !== 'false' && activo !== 'no' && activo !== '0';
    });
}

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
function _plMatch(s) {
  var n = (s||'').toLowerCase().replace(/[\s&_\-]+/g,'');
  // Coincide con: pl, fin-pl, operatingpl, p&l, pyg, etc.
  return n === 'pl' || n === 'operatingpl' || n === 'pyg' ||
         n === 'finpl' || n.indexOf('operatingpl') >= 0 ||
         /^.{0,6}pl$/.test(n); // termina en "pl" con hasta 6 chars de prefijo
}

function readViewData(ss, viewId, fechaInicio, fechaFin, sucursal, viewType, plMonth, plYear, plPrevYear) {
  sucursal   = sucursal   || 'Todas';
  viewType   = viewType   || '';
  plMonth    = plMonth    || '';
  plYear     = plYear     || '';
  plPrevYear = plPrevYear || '';
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

  // P&L check ANTES del fallback mensual — el fuente puede estar vacío en el menú
  if (_plMatch(fuente) || _plMatch(viewId)) {
    return readOperatingPL(viewType, plMonth, plYear, plPrevYear);
  }

  if (!fuente || fuente === 'mensual') {
    return readMensualData(ss, fechaInicio, fechaFin, viewId, sucursal);
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

  function _erMatch(s) {
    var n = (s || '').toLowerCase().replace(/\bde\b/g,'').replace(/[\s_\-]+/g,'');
    return n.indexOf('estado') > -1 && n.indexOf('result') > -1;
  }
  if (_erMatch(fuente) || _erMatch(viewId)) {
    return readEstadoResultados(fechaInicio, fechaFin);
  }

  if (fuente === 'Rep Ejecutivo' || viewId === 'rep-ejecutivo') {
    return readRepEjecutivo(ss, fechaInicio, fechaFin);
  }

  // Vistas de Reportes: devolver placeholder vacío hasta que se creen las hojas
  if (/^rep-/.test(viewId)) {
    return { view: viewId, fuente: fuente, rows: [], headers: [], periodo: fechaInicio + ' — ' + fechaFin };
  }

  // Cash Flow / Flujo de Efectivo — los datos reales vienen de action=banks
  if (/flujo|cashflow|cash.flow|efectivo|fin.cf/i.test(viewId) ||
      /flujo|cashflow|cash.flow|efectivo/i.test(fuente||'')) {
    return readBanksData();
  }

  return readCapturaData(ss, fuente, viewId, fechaInicio, fechaFin, sucursal);
}

/* ══════════════════════════════════════════════════════════════
   RESUMEN LABORATORIO — Dashboard clínico de calidad embrionaria
   Fuente: spreadsheet Lab (LAB_SS_ID)
   Hojas leídas: Resumen, ART Lab, FET, Inventario Crío, Insumos
   ══════════════════════════════════════════════════════════════ */