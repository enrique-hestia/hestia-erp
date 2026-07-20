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

/* ── Contraseñas: hash con sal (transición desde texto plano) ──────────────
   ANTES la hoja Usuarios guardaba la contraseña EN CLARO: cualquiera con lectura
   del libro las veía todas. Ahora se guardan como `h1$<sal>$<hash>` (SHA-256
   salado e iterado). _pwVerify acepta AMBOS formatos durante la transición, así
   que ninguna cuenta se bloquea: las contraseñas viejas (texto plano) siguen
   entrando y se convierten a hash al cambiarlas o al correr una vez
   hashearContrasenasExistentes(). Nota: esto es un ESCALÓN — el destino real es
   bcrypt/argon2 en Postgres (ver docs/ARQUITECTURA_ERP.md). NO se toca el token
   (usa sha256Hex+AUTH_SECRET, aparte). */
var _PW_VER  = 'h1';
var _PW_ITER = 300;                 // costo por verificación (login es poco frecuente)
function _pwSalt(){ return Utilities.getUuid().replace(/-/g,''); }   // 32 hex aleatorios
function _pwHash(pw, salt){
  var h = sha256Hex(salt + '|' + String(pw==null?'':pw));
  for (var i = 0; i < _PW_ITER; i++) h = sha256Hex(h + salt);
  return _PW_VER + '$' + salt + '$' + h;
}
function _pwIsHashed(stored){ return typeof stored === 'string' && stored.indexOf(_PW_VER + '$') === 0; }
function _pwVerify(stored, input){
  stored = String(stored == null ? '' : stored);
  input  = String(input  == null ? '' : input);
  if (stored === '') return false;              // vacío NUNCA pasa (fix previo)
  if (_pwIsHashed(stored)) {
    var parts = stored.split('$');              // h1$sal$hash
    if (parts.length !== 3) return false;
    return _pwHash(input, parts[1]) === stored; // recomputa con la MISMA sal
  }
  return stored === input;                       // legado: texto plano (transición)
}
/* Mantenimiento del dueño — convierte a hash las contraseñas que aún estén en
   texto plano. Idempotente (salta las ya hasheadas y las vacías). SEGURO 2×.
   Tras correrlo, ya no queda texto plano y el login sigue funcionando igual. */
function hashearContrasenasExistentes(){
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('Usuarios'); if (!sh) return {ok:false, error:'Hoja Usuarios no encontrada'};
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
    var pI = h.indexOf('contraseña'); if (pI < 0) return {ok:false, error:'No hay columna Contraseña'};
    var hasheadas=0, yaHash=0, vacias=0;
    for (var i=1;i<data.length;i++){
      var stored = String(data[i][pI]==null?'':data[i][pI]);
      if (stored === '') { vacias++; continue; }
      if (_pwIsHashed(stored)) { yaHash++; continue; }
      sh.getRange(i+1, pI+1).setValue(_pwHash(stored, _pwSalt()));
      hasheadas++;
    }
    return {ok:true, hasheadas:hasheadas, yaHasheadas:yaHash, vacias:vacias, total:data.length-1};
  } catch(ex){ return {ok:false, error:ex.message}; }
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
/* ── ALIAS DE USUARIO ─────────────────────────────────────────────────
   El Alias es SOLO capa de presentación (barra superior, chat). El Nombre
   completo sigue intacto en la hoja Usuarios porque lo usan nómina, CFDI y
   auditoría. Nunca devuelve vacío: alias → primer nombre → nombre → email.
   ─────────────────────────────────────────────────────────────────── */
function _aliasFor(alias, nombre, email) {
  var a = String(alias == null ? '' : alias).trim();
  if (a) return a;
  var n = String(nombre == null ? '' : nombre).trim();
  if (n) {
    var first = n.split(/\s+/)[0];
    if (first) return first;
    return n;
  }
  return String(email == null ? '' : email).trim();
}

// Devuelve la columna (1-indexed) cuyo header contiene `want` en la hoja
// Usuarios; si no existe la CREA al final (append-safe: nunca desplaza ni
// reordena columnas existentes). Mismo patrón que _egColEnsure/_ingColEnsure.
function _usrColEnsure(sh, want, headerText) {
  var lastCol = sh.getLastColumn();
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); });
  for (var c = 0; c < hdrs.length; c++) { if (hdrs[c].indexOf(want) > -1) return c + 1; }
  sh.getRange(1, lastCol + 1).setValue(headerText);
  sh.getRange(1, lastCol + 1).setFontWeight('bold').setBackground('#f3f4f6');
  return lastCol + 1;
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
  // ANTES: `if (user.password && user.password !== password)` — si la celda
  // Contraseña estaba VACÍA, el `user.password &&` cortocircuitaba y se aceptaba
  // CUALQUIER contraseña. Bastaba conocer el email de un admin sin contraseña
  // para entrar como admin. Ahora una contraseña vacía RECHAZA el login.
  if (!user.password)                    return { error: 'Tu usuario no tiene contraseña configurada. Contacta al administrador para que te asigne una.' };
  if (!_pwVerify(user.password, password)) return { error: 'Contraseña incorrecta.' };
  var rolCfg = getRolConfig(ss, user.rol);
  // Permisos operativos efectivos: admin/director = todo.
  var rl = String(user.rol||'').toLowerCase();
  var opPerms = rolCfg.permisosOperativos || [];
  if ((rl === 'admin' || rl === 'director') && opPerms.indexOf('*') === -1) opPerms = ['*'];
  var out = {
    success: true, token: generateToken(user.email),
    email: user.email, nombre: user.nombre, rol: user.rol,
    // alias: solo para mostrar en pantalla (barra superior / chat). El nombre
    // completo viaja igual en `nombre` para lo que ya dependa de él.
    alias: _aliasFor(user.alias, user.nombre, user.email),
    vistasBloqueadas: rolCfg.vistasBloqueadas,
    soloLectura:      rolCfg.soloLectura
  };
  // Solo mandar permisosOperativos si el rol YA está configurado; si no, el
  // frontend cae a su mapa por rol (no rompe a los usuarios existentes).
  if (opPerms.length) out.permisosOperativos = opPerms;
  return out;
}
function getUserRow(ss, email) {
  var sh = ss.getSheetByName('Usuarios');
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  var h    = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
  var eI=h.indexOf('email'), nI=h.indexOf('nombre'), pI=h.indexOf('contraseña'),
      rI=h.indexOf('rol'),   aI=h.indexOf('activo'), lI=h.indexOf('alias');
  for (var i=1; i<data.length; i++) {
    if (String(data[i][eI]).trim().toLowerCase() === email.toLowerCase()) {
      return { email: String(data[i][eI]).trim(), nombre: String(data[i][nI>-1?nI:0]).trim(),
               alias: lI>-1?String(data[i][lI]).trim():'',
               password: pI>-1?String(data[i][pI]).trim():'',
               rol: rI>-1?String(data[i][rI]).trim():'viewer',
               activo: aI>-1?data[i][aI]:true, rowNum: i+1 };
    }
  }
  return null;
}
function getRolConfig(ss, rol) {
  var sh  = ss.getSheetByName('Roles');
  var def = { vistasBloqueadas:[], soloLectura:false, permisosOperativos:[] };
  if (!sh) return def;
  var data = sh.getDataRange().getValues();
  var h    = data[0].map(function(c){ return String(c).trim().toLowerCase(); });
  var rI=h.indexOf('rol'), bI=h.indexOf('vistas_bloqueadas'), lI=h.indexOf('solo_lectura');
  var pI=h.indexOf('permisos_operativos'); if (pI<0) pI=h.indexOf('permisos');
  for (var i=1; i<data.length; i++) {
    if (String(data[i][rI]).trim().toLowerCase() === rol.toLowerCase()) {
      var bloq = bI>-1 ? String(data[i][bI]).split(',').map(function(v){return v.trim();}).filter(Boolean) : [];
      var ops  = pI>-1 ? String(data[i][pI]).split(',').map(function(v){return v.trim();}).filter(Boolean) : [];
      return { vistasBloqueadas: bloq, soloLectura: lI>-1 ? !!data[i][lI] : false, permisosOperativos: ops };
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

    /* ELIMINADOS 2026-07-18 por auditoría de seguridad — eran endpoints de debug
       SIN autenticación, servidos en una URL pública:
         · 'tabs'       — abría CUALQUIER spreadsheet cuyo id se pasara en ?sid=.
                          Como el Web App corre con la cuenta dueña, servía de
                          oráculo para sondear Nómina, Pacientes o Bancos.
         · 'labinspect' — devolvía un preview de TODAS las hojas del libro de
                          Laboratorio: resultados de pacientes en abierto.
       Ninguno lo usaba el frontend (0 referencias). No se re-agregan: si hace
       falta inspeccionar hojas, se hace desde el editor de Apps Script. */

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
      /* La clave estaba HARDCODEADA aquí y quedó en el historial de git →
         debe considerarse COMPROMETIDA. Ahora vive en una Script Property, que
         no viaja en el repo. Mientras no se defina, el endpoint queda CERRADO:
         así la clave filtrada deja de servir desde este mismo deploy.
         Para reactivarlo: Apps Script → Configuración → Propiedades del script →
         agregar `ECRAW_KEY` con una clave NUEVA (larga y aleatoria). */
      var ECRAW_KEY = '';
      try { ECRAW_KEY = PropertiesService.getScriptProperties().getProperty('ECRAW_KEY') || ''; } catch(ePK) {}
      if (!ECRAW_KEY) {
        return jsonResponse({ error: 'Endpoint deshabilitado: falta definir la Script Property ECRAW_KEY.', code: 403 });
      }
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

    // ── VALIDAR TOKEN en todas las acciones ──────────────────────────────────
    // ANTES esto era `if (shUsuariosExiste) { …validar… }`: si la hoja Usuarios
    // se renombraba, ocultaba o borraba, el candado se SALTABA ENTERO y las ~110
    // acciones quedaban públicas — y además _privResuelve() daba por bueno "ver
    // todo", apagando también el enmascarado de nombres de pacientes. Un cambio
    // accidental en una hoja de cálculo abría el ERP a internet.
    // Ahora falla CERRADO: sin módulo de usuarios no se atiende nada.
    var currentUser = null;
    var shUsuariosExiste = !!ss.getSheetByName('Usuarios');
    if (!shUsuariosExiste) {
      return jsonResponse({ error: 'Módulo de usuarios no disponible: la hoja Usuarios no existe. Por seguridad no se atienden solicitudes.', code: 503 });
    }
    {
      var tkn = (e && e.parameter.token) || '';
      var tkEmail = verifyToken(tkn);
      if (!tkEmail) return jsonResponse({ error: 'Sesión inválida. Inicia sesión nuevamente.', code: 401 });
      currentUser = getUserRow(ss, tkEmail);
      if (!currentUser || !currentUser.activo) return jsonResponse({ error: 'Usuario no autorizado.', code: 403 });
    }
    // Privacidad: fija si esta petición puede ver nombres sensibles (pacientes/empleados/proveedores).
    if (typeof _privSet === 'function') _privSet(ss, currentUser);

    // ── USUARIOS: listado para admin ──
    if (action === 'usuarios') {
      if (!currentUser || String(currentUser.rol||'').toLowerCase() !== 'admin')
        return jsonResponse({ error: 'Sin permisos de administrador.' });
      var shU2 = ss.getSheetByName('Usuarios');
      // Alias (presentación): se crea al final si aún no existe. Append-safe.
      try { _usrColEnsure(shU2, 'alias', 'Alias'); } catch (eAl) {}
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
      var opPerms    = (e && e.parameter.permisosOperativos);   // puede venir undefined
      if (!rolName) return jsonResponse({ error: 'Nombre de rol requerido.' });
      var shR  = ss.getSheetByName('Roles');
      if (!shR) return jsonResponse({ error: 'Hoja Roles no encontrada.' });
      var rData = shR.getDataRange().getValues();
      var rH    = rData[0].map(function(c){ return String(c).trim().toLowerCase(); });
      var rI  = rH.indexOf('rol');
      var bI  = rH.indexOf('vistas_bloqueadas');
      var lI  = rH.indexOf('solo_lectura');
      var dI  = rH.indexOf('descripcion');
      var oI  = rH.indexOf('permisos_operativos');
      // Si la columna de permisos operativos no existe, la creamos (para no perder el dato)
      if (oI < 0 && opPerms !== undefined && opPerms !== null) { oI = rData[0].length; shR.getRange(1, oI+1).setValue('permisos_operativos'); }
      var rowIdx = -1;
      for (var ri = 1; ri < rData.length; ri++) {
        if (String(rData[ri][rI]).trim().toLowerCase() === rolName.toLowerCase()) { rowIdx = ri + 1; break; }
      }
      if (rowIdx > 0) {
        if (bI > -1) shR.getRange(rowIdx, bI+1).setValue(bloqueadas);
        if (lI > -1) shR.getRange(rowIdx, lI+1).setValue(soloLec);
        if (dI > -1) shR.getRange(rowIdx, dI+1).setValue(desc);
        if (oI > -1 && opPerms !== undefined && opPerms !== null) shR.getRange(rowIdx, oI+1).setValue(opPerms);
      } else {
        var newRow = Array(Math.max(rData[0].length, oI+1)).fill('');
        if (rI > -1) newRow[rI] = rolName;
        if (bI > -1) newRow[bI] = bloqueadas;
        if (lI > -1) newRow[lI] = soloLec;
        if (dI > -1) newRow[dI] = desc;
        if (oI > -1 && opPerms !== undefined && opPerms !== null) newRow[oI] = opPerms;
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
        agencias:    (typeof _summaryAgencias === 'function' ? _summaryAgencias() : ['reprovida']),
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

    // Comisión MP actualmente registrada para un OP (para pre-cargar el modal al editar)
    if (action === 'mpComOp') {
      var _opMp = (e && (e.parameter.opId || e.parameter.op)) || '';
      var _cmp = (typeof _bankMpComisionDeOp === 'function' && _opMp) ? _bankMpComisionDeOp(_opMp) : 0;
      return jsonResponse({ ok:true, opId:_opMp, comision: Math.abs(_cmp) });
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

    // Tableros configurables ("Mi tablero"). El dueño se deriva del TOKEN, no de
    // un parámetro del cliente: nadie puede leer el tablero de otro. Fail-closed.
    if (action === 'dashboardLayouts') {
      if (typeof readDashboardLayouts !== 'function')
        return jsonResponse({ok:false, error:'Agrega dashboard.gs en Apps Script y redespliega.', layouts:[]});
      return jsonResponse(readDashboardLayouts(verifyToken((e && e.parameter.token) || '')));
    }

    // Catálogo de proveedores (fuente única para dropdowns)
    if (action === 'proveedores') {
      return jsonResponse(readProveedores());
    }
    // Datos fiscales del proveedor desde los XML recibidos (autollenado del alta)
    if (action === 'datosFiscalesProveedor') {
      if (typeof buscarDatosFiscalesProveedor !== 'function')
        return jsonResponse({ok:false, error:'Actualiza providers.gs + comprobantes.gs en Apps Script y redespliega.'});
      return jsonResponse(buscarDatosFiscalesProveedor((e && e.parameter.q) || ''));
    }

    // Presupuesto: proyección del siguiente trimestre + meta + pace
    if (action === 'presupuesto') {
      return jsonResponse(readPresupuesto((e && e.parameter.periodo) || ''));
    }
    // ── Comisiones por volumen de grupo (rebate a médicos externos + coordinador) ──
    // 'comisiones' SOLO CALCULA: no escribe un peso. Lo que se ve aquí es la pantalla
    // de revisión previa; escribir es 'generarComisiones' (POST, con permiso).
    if (action === 'comisiones') {
      if (typeof calcularComisiones !== 'function')
        return jsonResponse({ok:false, error:'Agrega comisiones.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(calcularComisiones({
        anio:    (e && e.parameter.anio)    || '',
        mes:     (e && e.parameter.mes)     || '',
        reglaId: (e && e.parameter.reglaId) || ''
      }));
    }
    if (action === 'comisionesCfg') {
      if (typeof readComisionesCfg !== 'function')
        return jsonResponse({ok:false, error:'Agrega comisiones.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readComisionesCfg());
    }
    // Presupuesto: histórico de metas por trimestre (racha meta vs real)
    if (action === 'historicoMetas') {
      if (typeof readHistoricoMetas !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(readHistoricoMetas());
    }
    // Presupuesto: grupos personalizados (config de agrupación)
    if (action === 'gruposPresupuesto') {
      if (typeof readGruposPresupuesto !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(readGruposPresupuesto());
    }
    if (action === 'presLocks') {
      if (typeof readPresLocks !== 'function')
        return jsonResponse({ok:true, locks:{}});
      return jsonResponse(readPresLocks());
    }
    // Presupuesto FP&A etapa 1: estructura de costos fijo/variable + P&L (utilidad)
    if (action === 'presupuestoCostos') {
      if (typeof readPresupuestoCostos !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(readPresupuestoCostos());
    }
    if (action === 'presupuestoModelo') {
      if (typeof readPresupuestoModelo !== 'function')
        return jsonResponse({ok:false, error:'Actualiza presupuesto.gs en Apps Script y redespliega.'});
      return jsonResponse(readPresupuestoModelo((e && e.parameter.periodo) || ''));
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
      var anioIng = (e && e.parameter.anio) || '';
      return jsonResponse(readIngresosData(anioIng ? parseInt(anioIng) : undefined));
    }

    if (action === 'origenes') {
      if (typeof readOrigenes !== 'function')
        return jsonResponse({ok:false, error:'Agrega origenes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readOrigenes());
    }

    if (action === 'sugerirOrigenes') {
      if (typeof sugerirOrigenesHistorico !== 'function')
        return jsonResponse({ok:false, error:'Agrega origenes.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(sugerirOrigenesHistorico((e && e.parameter.anio) || ''));
    }

    if (action === 'sugerirTiposProducto') {
      if (typeof sugerirTiposProducto !== 'function')
        return jsonResponse({ok:false, error:'Redespliega finance.gs.'});
      return jsonResponse(sugerirTiposProducto((e && e.parameter.anio) || ''));
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

    if (action === 'declaraciones') {
      if (typeof readDeclaraciones !== 'function') return jsonResponse({ ok: false, error: 'declaraciones.gs no desplegado' });
      var declPer = (e && e.parameter.periodo) || '';
      return jsonResponse(readDeclaraciones(declPer));
    }

    // ── Nómina ──
    // Todo lo de nómina (sueldos, RFC, NSS, bonos) exige el permiso 'editar_egresos'
    // —el MISMO que ya piden sus escrituras (nomina.gs)—, así que quien hace la
    // nómina ya lo tiene y no se rompe nada; admin/director pasan siempre. ANTES
    // estas lecturas solo pedían token: cualquier sesión (recepción, una vendedora)
    // se bajaba RFC+NSS+sueldos de toda la plantilla. EXCEPCIÓN: 'misRecibos' NO se
    // gatea — es autoservicio, ya filtrado a los recibos propios del empleado.
    var _NOMINA_GATED = { empleados:1, nominaMes:1, nominaCfg:1, nominaMesEstado:1, nominaBonos:1, nominaPeriodos:1, nominaCaptura:1, nominaSBC:1 };
    if (_NOMINA_GATED[action] && !_tokenHasPermission((e && e.parameter.token) || '', 'editar_egresos')) {
      return jsonResponse({ ok:false, error:'Sin autorización para ver nómina (editar_egresos). Pídeselo al administrador.', code:403 });
    }
    if (action === 'empleados') {
      if (typeof readEmpleados !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs no desplegado' });
      return jsonResponse(readEmpleados());
    }
    if (action === 'nominaMes') {
      if (typeof readNominaMes !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs no desplegado' });
      return jsonResponse(readNominaMes((e && e.parameter.anio) || '', (e && e.parameter.mes) || ''));
    }
    if (action === 'nominaCfg') {
      if (typeof nominaCfgRead !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs no desplegado' });
      return jsonResponse(nominaCfgRead());
    }
    if (action === 'misRecibos') {
      if (typeof nominaMisRecibos !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs no desplegado' });
      return jsonResponse(nominaMisRecibos((e && e.parameter.token) || '', (e && e.parameter.anio) || '', (e && e.parameter.mes) || ''));
    }
    if (action === 'nominaMesEstado') {
      if (typeof nominaMesEstado !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs no desplegado — redespliega core.gs/finance.gs/nomina.gs' });
      return jsonResponse(nominaMesEstado((e && e.parameter.anio) || '', (e && e.parameter.mes) || ''));
    }
    if (action === 'nominaBonos') {
      if (typeof readBonos !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs no desplegado — redespliega core.gs/finance.gs/nomina.gs' });
      return jsonResponse(readBonos((e && e.parameter.anio) || '', (e && e.parameter.mes) || ''));
    }
    // ── Nómina F5: periodos (semanal/quincenal/mensual), captura y SBC ──
    if (action === 'nominaPeriodos') {
      if (typeof readNominaPeriodos !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs (F5) no desplegado — redespliega core.gs/finance.gs/nomina.gs y corre setupNominaPeriodos()' });
      return jsonResponse(readNominaPeriodos((e && e.parameter.anio) || '', (e && e.parameter.tipo) || ''));
    }
    if (action === 'nominaCaptura') {
      if (typeof readNominaCaptura !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs (F5) no desplegado — redespliega core.gs/finance.gs/nomina.gs y corre setupNominaPeriodos()' });
      return jsonResponse(readNominaCaptura((e && e.parameter.periodoId) || ''));
    }
    if (action === 'nominaSBC') {
      if (typeof readNominaSBC !== 'function') return jsonResponse({ ok: false, error: 'nomina.gs (F5) no desplegado — redespliega core.gs/finance.gs/nomina.gs y corre setupNominaPeriodos()' });
      return jsonResponse(readNominaSBC((e && e.parameter.anio) || ''));
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

    if (action === 'abonosOrden') {
      if (typeof readAbonosOrden !== 'function')
        return jsonResponse({ok:false, error:'Agrega cxp_creditos.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readAbonosOrden((e && e.parameter.cxpId) || ''));
    }

    if (action === 'resumenSemanal') {
      if (typeof readResumenSemanal !== 'function')
        return jsonResponse({ok:false, error:'Agrega semanal.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readResumenSemanal((e && e.parameter.fecha) || ''));
    }
    if (action === 'semanalConfig') {
      if (typeof readSemanalConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega semanal.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse({ok:true, config: readSemanalConfig()});
    }

    if (action === 'summary') {
      if (typeof readSummary !== 'function')
        return jsonResponse({ok:false, error:'Agrega summary.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readSummary((e && e.parameter.fechaInicio) || '', (e && e.parameter.fechaFin) || ''));
    }

    if (action === 'estadoResultadosMensual') {
      if (typeof readEstadoResultadosMensual !== 'function')
        return jsonResponse({ok:false, error:'Actualiza summary.gs en Apps Script y redespliega.'});
      return jsonResponse(readEstadoResultadosMensual((e && e.parameter.fechaInicio) || '', (e && e.parameter.fechaFin) || ''));
    }
    if (action === 'resumenGeneral') {
      if (typeof readResumenGeneral !== 'function')
        return jsonResponse({ok:false, error:'Actualiza summary.gs en Apps Script y redespliega.'});
      return jsonResponse(readResumenGeneral((e && e.parameter.fechaInicio) || '', (e && e.parameter.fechaFin) || ''));
    }

    if (action === 'summaryConfig') {
      if (typeof readSummaryConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega summary.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readSummaryConfig());
    }

    if (action === 'gastoDevengado') {
      if (typeof readGastoDevengado !== 'function')
        return jsonResponse({ok:false, error:'Agrega devengado.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readGastoDevengado((e && e.parameter.fechaInicio) || '', (e && e.parameter.fechaFin) || ''));
    }
    if (action === 'factVencida') {
      if (typeof factVencidaLeer !== 'function')
        return jsonResponse({ok:false, error:'Agrega devengado.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(factVencidaLeer());
    }
    if (action === 'agencias') {
      if (typeof agenciasLeer !== 'function')
        return jsonResponse({ok:false, error:'Actualiza summary.gs en Apps Script y redespliega.'});
      return jsonResponse(agenciasLeer());
    }

    if (action === 'boardReport') {
      if (typeof readBoardReport !== 'function')
        return jsonResponse({ok:false, error:'Agrega board.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readBoardReport((e && e.parameter.fechaInicio) || '', (e && e.parameter.fechaFin) || ''));
    }
    if (action === 'boardConfig') {
      if (typeof readBoardConfig !== 'function')
        return jsonResponse({ok:false, error:'Agrega board.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readBoardConfig());
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

    /* ── IDENTIFICAR PACIENTES ──
       Cobros "No localizado" pendientes de nombre. Va por GET porque la
       respuesta NO lleva datos personales (solo OP, fecha, monto y método);
       el nombre de la paciente viaja siempre por POST, nunca en la URL. */
    if (action === 'noLocalizados') {
      if (typeof listarNoLocalizados !== 'function')
        return jsonResponse({ ok:false, error:'Agrega identificar.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(listarNoLocalizados({
        token:     (e && e.parameter.token) || '',
        anio:      (e && e.parameter.anio) || '',
        fechaIni:  (e && e.parameter.fechaIni) || '',
        fechaFin:  (e && e.parameter.fechaFin) || '',
        montoMin:  (e && e.parameter.montoMin) || '',
        montoMax:  (e && e.parameter.montoMax) || ''
      }));
    }

    // ── COBRANZA (Cuentas por Cobrar) — Motor A (saldos) + Motor B (suscripciones crío) ──
    if (action === 'cuentasCobrar') {
      if (typeof readCuentasPorCobrar !== 'function')
        return jsonResponse({ ok:false, error:'Agrega cobranza.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(readCuentasPorCobrar({}));
    }
    if (action === 'suscripcionesCrio') {
      if (typeof readSuscripcionesCrio !== 'function')
        return jsonResponse({ ok:false, error:'Agrega cobranza.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(readSuscripcionesCrio({}));
    }
    if (action === 'estadoCobranza') {
      if (typeof readEstadoCobranza !== 'function')
        return jsonResponse({ ok:false, error:'Agrega cobranza.gs al proyecto de Apps Script y redespliega.' });
      var pCob = e && e.parameter.paciente ? decodeURIComponent(e.parameter.paciente) : '';
      return jsonResponse(readEstadoCobranza(pCob));
    }
    if (action === 'cobranzaSetup') {
      if (typeof setupCobranzaConfig !== 'function')
        return jsonResponse({ ok:false, error:'Agrega cobranza.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(setupCobranzaConfig());
    }
    // Traducción de textos dinámicos (descripciones de servicio) para la Carta de Seguro.
    if (action === 'traducir') {
      if (typeof traducirTextos !== 'function')
        return jsonResponse({ ok:false, error:'Agrega cobranza.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(traducirTextos({
        lang:   (e && e.parameter.lang) || 'en',
        textos: (e && e.parameter.textos) || '[]'
      }));
    }
    /* ── TRADUCCIONES (catálogo = única fuente de verdad) ──────────────────
       tradMap: lo que consulta el frontend ANTES de imprimir para decidir si
       bloquea. Devuelve { normalizado(Descripcion) → traducción }. */
    if (action === 'tradMap') {
      if (typeof readTradMap !== 'function')
        return jsonResponse({ ok:false, error:'Agrega traducciones.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(readTradMap((e && e.parameter.lang) || 'en'));
    }
    if (action === 'tradAuditar') {
      if (typeof auditarTraduccionesProductos !== 'function')
        return jsonResponse({ ok:false, error:'Agrega traducciones.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(auditarTraduccionesProductos({ lang: (e && e.parameter.lang) || '' }));
    }
    if (action === 'paisIdioma') {
      if (typeof readPaisIdiomaMap !== 'function')
        return jsonResponse({ ok:false, error:'Agrega traducciones.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(readPaisIdiomaMap());
    }
    // Cobranza: configuración de descuentos por agencia (panel)
    if (action === 'cobDescCfg') {
      if (typeof cobDescCfgRead !== 'function')
        return jsonResponse({ ok:false, error:'Actualiza cobranza.gs en Apps Script y redespliega.' });
      return jsonResponse(cobDescCfgRead());
    }
    if (action === 'generarSuscripcionesPreview') {
      if (typeof generarSuscripciones !== 'function')
        return jsonResponse({ ok:false, error:'Agrega cobranza.gs al proyecto de Apps Script y redespliega.' });
      return jsonResponse(generarSuscripciones({ preview:true }));
    }
    // Conciliación Mercado Pago (crío): estado actual usando SOLO las ligas persistidas.
    if (action === 'conciliarCrioEstado') {
      if (typeof conciliarSuscripcionesMP !== 'function')
        return jsonResponse({ ok:false, error:'Actualiza cobranza.gs en Apps Script y redespliega.' });
      return jsonResponse(conciliarSuscripcionesMP({ soloEstado:true, token: (e && e.parameter.token) || '' }));
    }
    if (action === 'contarEgresosProveedor') {
      if (typeof contarEgresosProveedor !== 'function')
        return jsonResponse({ ok:false, error:'Actualiza providers.gs en Apps Script y redespliega.' });
      var pnom = e && e.parameter.nombre ? decodeURIComponent(e.parameter.nombre) : '';
      return jsonResponse(contarEgresosProveedor(pnom));
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

    if (action === 'getDropdowns') {
      if (typeof readDropdowns !== 'function')
        return jsonResponse({ok:false, error:'Falta readDropdowns() — actualiza finance.gs en Apps Script y redespliega.'});
      return jsonResponse(readDropdowns());
    }
    // Ruteo de bancos (formas de pago → banco + comisión) + catálogo de cuentas,
    // para que el frontend decida banco/comisión desde la config (no hardcode).
    if (action === 'rutasBancos') {
      if (typeof _formasPago !== 'function')
        return jsonResponse({ok:false, error:'Falta _formasPago() — actualiza finance.gs y redespliega.'});
      return jsonResponse({ok:true, formas:_formasPago(), cuentas:_cuentas()});
    }

    if (action === 'saveDropdown') {
      if (typeof saveDropdownValues !== 'function')
        return jsonResponse({ok:false, error:'Falta saveDropdownValues() — actualiza finance.gs en Apps Script y redespliega.'});
      var ddData = e.parameter.data || '{}';
      return jsonResponse(saveDropdownValues(JSON.parse(decodeURIComponent(ddData))));
    }

    // ── Config del aviso de Novedades (novedades.gs) ──
    // Lectura abierta a cualquier sesión: el filtro de audiencia corre en el frontend
    // con hasPermission y para eso todos deben poder leer las banderas.
    if (action === 'getNovedadesCfg') {
      if (typeof readNovedadesCfg !== 'function')
        return jsonResponse({ok:false, error:'Falta readNovedadesCfg() — agrega novedades.gs al proyecto de Apps Script y redespliega.'});
      return jsonResponse(readNovedadesCfg());
    }
    // El guardado valida admin/director DENTRO de saveNovedadesCfg (server-side).
    if (action === 'saveNovedadesCfg') {
      if (typeof saveNovedadesCfg !== 'function')
        return jsonResponse({ok:false, error:'Falta saveNovedadesCfg() — agrega novedades.gs al proyecto de Apps Script y redespliega.'});
      var novData = e.parameter.data || '{}';
      return jsonResponse(saveNovedadesCfg(JSON.parse(decodeURIComponent(novData))));
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
        // ── Canal y Médico Tratante YA NO salen de la pestaña "Opciones":
        //    salen del CATÁLOGO Origenes_Externos (origenes.gs), que es el
        //    mismo que atribuye las ventas. Antes eran texto libre paralelo y
        //    por eso no se podía reportar por médico/agencia.
        //      Canal  → orígenes tipo Grupo y Agencia ("de dónde viene").
        //      Médico → orígenes tipo "Médico externo".
        //    `origenes` viaja para que el frontend filtre los médicos por el
        //    grupo del Canal elegido.
        //    Si el catálogo está vacío o el módulo no está desplegado,
        //    _origOpcionesFicha() devuelve ok:false y se conservan tal cual
        //    las opciones de la hoja (nunca se deja la ficha sin opciones).
        if (String(sheetName) === 'Pacientes' && typeof _origOpcionesFicha === 'function') {
          var _of = _origOpcionesFicha();
          if (_of && _of.ok) {
            options['Canal']           = _of.canales;
            options['Médico Tratante'] = _of.medicos;
            return jsonResponse({ options: options, origenes: _of.registros, origenesFicha: true });
          }
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
        // Duplicado en Pacientes al EDITAR: correo/nombre ya en OTRO paciente
        // (se excluye la propia fila). Cierra la puerta trasera del alta.
        if (sheetName.trim().toLowerCase() === 'pacientes') {
          var uNombreIdx=-1, uEmailIdx=-1, uIdIdx=-1;
          for (var uh=0; uh<hdrs.length; uh++){
            var uhl=hdrs[uh].toLowerCase();
            if(uNombreIdx<0 && uhl.indexOf('nombre')>-1) uNombreIdx=uh;
            if(uEmailIdx<0 && (uhl.indexOf('mail')>-1||uhl==='correo')) uEmailIdx=uh;
            if(uIdIdx<0 && uhl==='id') uIdIdx=uh;
          }
          var uNom   = uNombreIdx>-1 ? String(e.parameter[hdrs[uNombreIdx]]||'').trim().toLowerCase() : '';
          var uEmail = uEmailIdx >-1 ? String(e.parameter[hdrs[uEmailIdx]] ||'').trim().toLowerCase() : '';
          if (uNom || uEmail) {
            var uAll = shUpd.getDataRange().getValues();
            for (var ur=hdrInfo.dataStart; ur<uAll.length; ur++){
              if (ur === rowNum-1) continue; // la propia fila que se edita
              if (uNom && String(uAll[ur][uNombreIdx]||'').trim().toLowerCase()===uNom)
                return jsonResponse({ error:'Ya existe otro paciente con el nombre "'+e.parameter[hdrs[uNombreIdx]]+'".', duplicado:true });
              if (uEmail && String(uAll[ur][uEmailIdx]||'').trim().toLowerCase()===uEmail){
                var uqId=uIdIdx>-1?String(uAll[ur][uIdIdx]||''):'', uqNom=uNombreIdx>-1?String(uAll[ur][uNombreIdx]||''):'';
                return jsonResponse({ error:'El correo "'+e.parameter[hdrs[uEmailIdx]]+'" ya está registrado'+((uqId||uqNom)?' en el paciente '+(uqId?uqId+' — ':'')+uqNom:'')+'.', duplicado:true, duplicadoCorreo:true });
              }
            }
          }
        }
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
      try { CacheService.getScriptCache().removeAll(['erp_menu_v2_' + fechaInicio + '_' + fechaFin, 'erp_menu_v3_' + fechaInicio + '_' + fechaFin, 'erp_banks_v1']); } catch(e) {}
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

/* Actualiza SOLO los campos editables del menú (padre, label, icono, orden, activo)
   por id, in-place. NUNCA toca id/tipo/fuente/seccion para no romper el ruteo, y no
   borra filas que no vengan en el body (evita perder items por accidente). */
function saveMenu(body) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('Menu');
    if (!sh) return { ok:false, error:'No existe la hoja Menu' };
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return { ok:false, error:'Hoja Menu vacía' };
    var ncol = vals[0].length;
    var idRow = {};
    for (var i = 1; i < vals.length; i++) { var id = String(vals[i][0]||'').trim(); if (id) idRow[id] = i; }
    var updated = 0, creados = 0;
    (body.items || []).forEach(function(it){
      var id = String(it.id||'').trim(); if (!id) return;
      if (id in idRow) {
        var i = idRow[id];
        if (it.padre  !== undefined) vals[i][1] = String(it.padre||'');
        if (it.label  !== undefined) vals[i][2] = String(it.label||'');
        if (it.icono  !== undefined) vals[i][4] = String(it.icono||'');
        if (it.orden  !== undefined) vals[i][5] = Number(it.orden)||0;
        if (it.activo !== undefined) vals[i][8] = !(it.activo === false || String(it.activo).toLowerCase() === 'false');
        updated++;
      } else {
        // Item NUEVO (submenú creado desde la página): se agrega con sus campos.
        var row = new Array(ncol);
        row[0] = id;
        row[1] = String(it.padre||'');
        row[2] = String(it.label||'');
        row[3] = String(it.seccion||'');
        row[4] = String(it.icono||'circle');
        row[5] = Number(it.orden)||0;
        row[6] = String(it.tipo||'vista');
        row[7] = String(it.fuente||'');
        if (ncol > 8) row[8] = !(it.activo === false || String(it.activo).toLowerCase() === 'false');
        vals.push(row); idRow[id] = vals.length-1; creados++;
      }
    });
    sh.getRange(1, 1, vals.length, ncol).setValues(vals);
    try { logAudit(body.usuario||'sistema', 'Panel', 'Guardar menú', '', '', '', updated + ' items, ' + creados + ' nuevos'); } catch(e) {}
    return { ok:true, updated:updated, creados:creados };
  } catch (ex) { return { ok:false, error:ex.message }; }
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
    var _pl = readOperatingPL(viewType, plMonth, plYear, plPrevYear);
    // Conecta el módulo Presupuesto como fuente del Budget (reemplaza la pestaña manual)
    if (typeof _presInyectaBudgetEnPL === 'function') { try { _pl = _presInyectaBudgetEnPL(_pl, viewType, plMonth, plYear); } catch(e){} }
    return _pl;
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