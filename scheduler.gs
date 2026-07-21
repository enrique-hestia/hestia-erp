/* ==============================================================
   scheduler.gs — Tareas programadas y caché de configuración
   --------------------------------------------------------------
   Corre diariamente a las 23:00 hrs y actualiza toda la
   configuración del sistema en Script Properties (caché rápida)
   para que las peticiones HTTP del día siguiente respondan
   sin leer Google Sheets en cada llamada.

   Flujo:
     1. actualizarConfiguracionSistema()  ← disparado por trigger
        a. Lee hoja "Configuracion" (SHEET_ID) → Script Properties
        b. Lee hoja "IMSS_Parametros" (SHEET_ID) → Script Properties
        c. Refresca caché de menú y usuarios
        d. Valida conectividad de todos los spreadsheets clave
        e. Escribe log en hoja "Config_Log"

   Helpers públicos (usables desde cualquier módulo):
     getParam(clave, valorDefault)  — lee un parámetro del sistema
     getParamNum(clave, default)    — versión numérica
     getParamBool(clave, default)   — versión booleana
     clearSystemCache()             — limpia caché de CacheService

   Setup (correr UNA sola vez desde el editor de Apps Script):
     setupNightlyTrigger()          — instala el trigger diario 23:00
     removeNightlyTrigger()         — lo elimina
   ============================================================== */

/* ── Prefijos de cache ───────────────────────────────────────── */
var CFG_PREFIX  = 'cfg_';
var CFG_VERSION = 'cfg_version';
var CFG_LAST_RUN = 'cfg_last_run';

/* ══════════════════════════════════════════════════════════════
   INSTALACIÓN DEL TRIGGER — correr UNA vez desde el editor
   ══════════════════════════════════════════════════════════════ */

function setupNightlyTrigger() {
  // Eliminar triggers anteriores del mismo nombre para evitar duplicados
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'actualizarConfiguracionSistema') {
      ScriptApp.deleteTrigger(existing[i]);
    }
  }
  // Crear trigger diario a las 23:00 hrs (hora del servidor / zona Config)
  ScriptApp.newTrigger('actualizarConfiguracionSistema')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(0)
    .create();
  Logger.log('[scheduler] Trigger nocturno instalado: actualizarConfiguracionSistema @ 23:00 cada día.');
  return 'Trigger instalado correctamente.';
}

function removeNightlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed  = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'actualizarConfiguracionSistema') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('[scheduler] Triggers eliminados: ' + removed);
  return removed + ' trigger(s) eliminado(s).';
}

/* ══════════════════════════════════════════════════════════════
   FUNCIÓN PRINCIPAL — se ejecuta cada noche a las 23:00
   ══════════════════════════════════════════════════════════════ */

function actualizarConfiguracionSistema() {
  var inicio   = new Date();
  var log      = [];
  var errores  = [];
  var props    = PropertiesService.getScriptProperties();
  var cache    = CacheService.getScriptCache();

  _logMsg(log, 'INFO', 'Inicio actualización de configuración del sistema');

  // ── 1. Leer hoja "Configuracion" → Script Properties ──────────
  try {
    var n1 = _actualizarParametrosGenerales(props, log);
    _logMsg(log, 'OK', 'Parámetros generales actualizados: ' + n1 + ' clave(s)');
  } catch(e) {
    _logMsg(log, 'ERROR', 'Parámetros generales: ' + e.message);
    errores.push('Parametros: ' + e.message);
  }

  // ── 2. Leer hoja "IMSS_Parametros" → Script Properties ────────
  try {
    var n2 = _actualizarIMSS(props, log);
    _logMsg(log, 'OK', 'Parámetros IMSS actualizados: ' + n2 + ' clave(s)');
  } catch(e) {
    _logMsg(log, 'WARN', 'IMSS_Parametros no disponible: ' + e.message);
  }

  // ── 3. Actualizar tabla de usuarios en Script Properties ───────
  try {
    var n3 = _actualizarUsuariosCache(props, log);
    _logMsg(log, 'OK', 'Caché de usuarios actualizada: ' + n3 + ' usuario(s)');
  } catch(e) {
    _logMsg(log, 'ERROR', 'Usuarios: ' + e.message);
    errores.push('Usuarios: ' + e.message);
  }

  // ── 4. Limpiar caché de CacheService (menú, bancos, etc.) ──────
  try {
    _limpiarCacheService(cache, log);
    _logMsg(log, 'OK', 'CacheService limpiado — se reconstruirá en la primera petición del día');
  } catch(e) {
    _logMsg(log, 'WARN', 'CacheService: ' + e.message);
  }

  // ── 5. Validar conectividad de todos los spreadsheets críticos ─
  try {
    var ssCheck = _validarConectividad(log);
    _logMsg(log, 'OK', 'Conectividad validada: ' + ssCheck.ok + ' ok, ' + ssCheck.fail + ' falló');
    if (ssCheck.fail > 0) errores.push('Conectividad: ' + ssCheck.errores.join(', '));
  } catch(e) {
    _logMsg(log, 'ERROR', 'Validación conectividad: ' + e.message);
  }

  // ── 6. Actualizar período activo y versión de la API ──────────
  try {
    _actualizarPeriodosActivos(props, log);
    _logMsg(log, 'OK', 'Períodos y versión de API actualizados');
  } catch(e) {
    _logMsg(log, 'WARN', 'Períodos: ' + e.message);
  }

  // ── 7. Registrar en Script Properties la hora del último run ──
  var fin = new Date();
  var duracion = ((fin - inicio) / 1000).toFixed(1);
  props.setProperty(CFG_LAST_RUN, fin.toISOString());
  props.setProperty(CFG_VERSION,  API_VERSION + '|' + fin.toISOString());
  _logMsg(log, 'INFO', 'Actualización completada en ' + duracion + 's. Errores: ' + errores.length);

  // ── 8. Escribir log en hoja "Config_Log" ─────────────────────
  try {
    _escribirLog(log, fin, duracion, errores.length);
  } catch(e) {
    Logger.log('[scheduler] No se pudo escribir log: ' + e.message);
  }

  Logger.log('[scheduler] actualizarConfiguracionSistema finalizado en ' + duracion + 's');
  return { ok: true, duracion: duracion + 's', errores: errores, lineas: log.length };
}

/* ══════════════════════════════════════════════════════════════
   SUBFUNCIONES DE ACTUALIZACIÓN
   ══════════════════════════════════════════════════════════════ */

/* Lee hoja "Configuracion" (columnas: Clave | Valor | Descripcion | Activo)
   y guarda cada clave activa en Script Properties con prefijo cfg_         */
function _actualizarParametrosGenerales(props, log) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var sh  = ss.getSheetByName('Configuracion');
  if (!sh) {
    _setupHojaConfiguracion(ss); // la crea si no existe
    sh = ss.getSheetByName('Configuracion');
    _logMsg(log, 'INFO', 'Hoja "Configuracion" creada con valores default');
  }
  var raw = sh.getDataRange().getValues();
  if (raw.length < 2) return 0;
  var H   = raw[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iC  = H.indexOf('clave');
  var iV  = H.indexOf('valor');
  var iA  = H.indexOf('activo');
  if (iC < 0 || iV < 0) throw new Error('Hoja Configuracion sin columnas Clave/Valor');
  var count = 0;
  for (var r = 1; r < raw.length; r++) {
    var clave  = String(raw[r][iC] || '').trim();
    var valor  = String(raw[r][iV] || '').trim();
    var activo = iA < 0 || raw[r][iA] === '' || raw[r][iA] === true || String(raw[r][iA]).toLowerCase() === 'true' || String(raw[r][iA]) === '1';
    if (!clave) continue;
    if (activo) {
      props.setProperty(CFG_PREFIX + clave, valor);
      count++;
    } else {
      try { props.deleteProperty(CFG_PREFIX + clave); } catch(e) {}
    }
    // Registrar la última actualización en la misma hoja (col E)
    try { sh.getRange(r+1, 5).setValue(new Date()); } catch(e) {}
  }
  return count;
}

/* Lee hoja "IMSS_Parametros" con tasas/UMA/salarios mínimos          */
function _actualizarIMSS(props, log) {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var sh  = ss.getSheetByName('IMSS_Parametros');
  if (!sh) return 0;
  var raw = sh.getDataRange().getValues();
  if (raw.length < 2) return 0;
  var H  = raw[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iC = H.indexOf('clave') > -1 ? H.indexOf('clave') : 0;
  var iV = H.indexOf('valor') > -1 ? H.indexOf('valor') : 1;
  var count = 0;
  for (var r = 1; r < raw.length; r++) {
    var clave = String(raw[r][iC] || '').trim();
    var valor = String(raw[r][iV] || '').trim();
    if (!clave || !valor) continue;
    props.setProperty(CFG_PREFIX + 'imss_' + clave, valor);
    count++;
  }
  return count;
}

/* Serializa emails y roles de la hoja Usuarios en Script Properties  */
function _actualizarUsuariosCache(props, log) {
  var ss   = SpreadsheetApp.openById(SHEET_ID);
  var sh   = ss.getSheetByName('Usuarios');
  if (!sh) return 0;
  var raw  = sh.getDataRange().getValues();
  if (raw.length < 2) return 0;
  var H    = raw[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iE   = H.indexOf('email');
  var iR   = H.indexOf('rol');
  var iA   = H.indexOf('activo');
  if (iE < 0) return 0;
  var usuarios = [];
  for (var r = 1; r < raw.length; r++) {
    var email  = String(raw[r][iE] || '').trim().toLowerCase();
    var rol    = iR > -1 ? String(raw[r][iR] || 'viewer').trim() : 'viewer';
    var activo = iA < 0 || raw[r][iA] === '' || raw[r][iA] === true || String(raw[r][iA]).toLowerCase() === 'true';
    if (!email) continue;
    if (activo) usuarios.push(email + ':' + rol);
  }
  // Guardar como string compacto; máx ~9KB por propiedad
  props.setProperty(CFG_PREFIX + '_usuarios', usuarios.join('|'));
  return usuarios.length;
}

/* Borra las entradas de CacheService conocidas para forzar refresco  */
function _limpiarCacheService(cache, log) {
  var keysToDelete = [
    'erp_menu_v2',
    'erp_banks_v1',
    'erp_menu_v1',
    'erp_menu_v3'
  ];
  // Borrar claves exactas conocidas
  cache.removeAll(keysToDelete);
  // CacheService no permite listar todas las claves, pero las rutas de menú
  // incluyen la fecha, así que también regeneramos para hoy y mañana
  var today    = _fmtDateLocal(new Date());
  var tomorrow = _fmtDateLocal(new Date(Date.now() + 86400000));
  var extraKeys = [];
  [today, tomorrow].forEach(function(d) {
    extraKeys.push('erp_menu_v2_' + d + '_01_' + d + '_31');
  });
  try { cache.removeAll(extraKeys); } catch(e) {}
  _logMsg(log, 'INFO', 'CacheService: claves limpiadas');
}

/* Abre cada spreadsheet clave y verifica que tenga hojas              */
function _validarConectividad(log) {
  var IDS = {
    'ERP Principal':     SHEET_ID,
    'Estado Resultados': ER_SS_ID,
    'Bancos':            BANKS_SS_ID,
    'Laboratorio':       LAB_SS_ID,
    'Medicamentos':      MED_SS_ID,
    'Pacientes':         PAC_SS_ID,
    'Productos':         PROD_SS_ID,
    'Caja Chica':        CAJA_CHICA_SS_ID,
    'CxP':               CXP_SS_ID
  };
  var ok = 0, fail = 0, erroresV = [];
  Object.keys(IDS).forEach(function(nombre) {
    try {
      var ss   = SpreadsheetApp.openById(IDS[nombre]);
      var tabs = ss.getSheets().length;
      _logMsg(log, 'OK', nombre + ': ' + tabs + ' pestaña(s)');
      ok++;
    } catch(e) {
      _logMsg(log, 'ERROR', nombre + ' [' + IDS[nombre] + ']: ' + e.message);
      erroresV.push(nombre);
      fail++;
    }
  });
  return { ok: ok, fail: fail, errores: erroresV };
}

/* Calcula y guarda el período activo (año fiscal, mes actual, etc.)   */
function _actualizarPeriodosActivos(props, log) {
  var hoy    = new Date();
  var anio   = hoy.getFullYear();
  var mes    = String(hoy.getMonth() + 1).padStart(2, '0');
  var dia    = String(hoy.getDate()).padStart(2, '0');
  var mesNom = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'][hoy.getMonth()];
  // Determinar trimestre
  var trimestre = 'Q' + (Math.ceil((hoy.getMonth() + 1) / 3));
  var periodoId = anio + '-' + trimestre;
  // Días hábiles aproximados del mes (sin weekends)
  var diasHabiles = 0;
  var lastDay     = new Date(anio, hoy.getMonth() + 1, 0).getDate();
  for (var d = 1; d <= lastDay; d++) {
    var dw = new Date(anio, hoy.getMonth(), d).getDay();
    if (dw !== 0 && dw !== 6) diasHabiles++;
  }
  // Guardar
  props.setProperty(CFG_PREFIX + 'anio_actual',      String(anio));
  props.setProperty(CFG_PREFIX + 'mes_actual',        mes);
  props.setProperty(CFG_PREFIX + 'mes_nombre',        mesNom);
  props.setProperty(CFG_PREFIX + 'dia_actual',        dia);
  props.setProperty(CFG_PREFIX + 'trimestre_actual',  trimestre);
  props.setProperty(CFG_PREFIX + 'periodo_id',        periodoId);
  props.setProperty(CFG_PREFIX + 'dias_habiles_mes',  String(diasHabiles));
  props.setProperty(CFG_PREFIX + 'api_version',       API_VERSION);
  _logMsg(log, 'INFO', 'Período activo: ' + periodoId + ' (' + mesNom + ' ' + anio + ', ' + diasHabiles + ' días hábiles)');
}

/* ══════════════════════════════════════════════════════════════
   LOG Y ESCRITURA EN HOJA Config_Log
   ══════════════════════════════════════════════════════════════ */

function _logMsg(log, nivel, msg) {
  log.push({ t: new Date(), nivel: nivel, msg: msg });
  Logger.log('[scheduler][' + nivel + '] ' + msg);
}

function _escribirLog(log, fin, duracion, numErrores) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sh    = ss.getSheetByName('Config_Log');
  if (!sh) {
    sh = ss.insertSheet('Config_Log');
    sh.getRange(1,1,1,5).setValues([['Timestamp','Nivel','Mensaje','Duración(s)','Errores']]);
    sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#c46a7a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 5, [160, 70, 560, 90, 60]);
  }
  // Escritura masiva en una sola llamada
  var rows = log.map(function(l) {
    return [_isoLocal(l.t), l.nivel, l.msg, '', ''];
  });
  // Última fila: resumen
  rows.push([_isoLocal(fin), 'RESUMEN', 'Ciclo completado', duracion, numErrores]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  // Colorear según nivel
  var firstRow = sh.getLastRow() - rows.length + 1;
  rows.forEach(function(r, i) {
    var nivel = r[1];
    var color = nivel === 'OK'?'#d1fae5': nivel === 'ERROR'?'#fee2e2': nivel === 'WARN'?'#fef3c7':'#f9fafb';
    sh.getRange(firstRow + i, 1, 1, 5).setBackground(color);
  });
  // Mantener solo las últimas 500 filas de log (evitar crecer sin límite)
  var maxRows  = 501; // 500 + encabezado
  var totalRows = sh.getLastRow();
  if (totalRows > maxRows) {
    sh.deleteRows(2, totalRows - maxRows);
  }
}

/* ══════════════════════════════════════════════════════════════
   SETUP: CREAR HOJA "Configuracion" CON PARÁMETROS DEFAULT
   ══════════════════════════════════════════════════════════════ */

function _setupHojaConfiguracion(ss) {
  var sh = ss.insertSheet('Configuracion');
  var HDRS = ['Clave', 'Valor', 'Descripcion', 'Activo', 'Ultima_Actualizacion'];
  var DEFAULTS = [
    // Clínica
    ['clinica_nombre',           'Hestia Fertility',                    'Nombre oficial de la clínica',                    true,  ''],
    ['clinica_ciudad',           'Guadalajara',                         'Ciudad de operación principal',                   true,  ''],
    ['clinica_moneda',           'MXN',                                 'Moneda base (MXN | USD)',                         true,  ''],
    ['clinica_tipo_cambio',      '17.50',                               'Tipo de cambio USD→MXN (actualizar manualmente)', true,  ''],
    // Fiscal
    ['fiscal_anio_inicio',       '01-01',                               'Inicio del año fiscal (MM-DD)',                  true,  ''],
    ['fiscal_meses_proyeccion',  '3',                                   'Meses de proyección de tendencia',                true,  ''],
    // Alertas / KPIs
    ['kpi_ocupacion_umbral',     '75',                                  'Porcentaje mínimo de ocupación esperado',         true,  ''],
    ['kpi_margen_umbral',        '30',                                  'Margen operativo mínimo esperado (%)',            true,  ''],
    ['kpi_budget_semaforo_verde','100',                                  'Cumplimiento budget para semáforo verde (%)',     true,  ''],
    ['kpi_budget_semaforo_amarillo','90',                               'Cumplimiento budget para semáforo amarillo (%)',  true,  ''],
    // Presupuesto
    ['pres_meses_proyeccion',    '4',                                   'Meses futuros en gráfica de presupuesto',         true,  ''],
    ['pres_factor_crec',         '1.05',                                'Factor de crecimiento anual esperado',            true,  ''],
    // Ingresos
    ['ing_ticket_objetivo_mxn',  '80000',                               'Ticket promedio objetivo por paciente (MXN)',     true,  ''],
    ['ing_nuevos_pac_mes',       '15',                                  'Meta de nuevos pacientes por mes',                true,  ''],
    // Surrogacy
    ['surrogacy_ticket_usd',     '85000',                               'Tarifa base programa Surrogacy (USD)',             true,  ''],
    ['surrogacy_casos_meta_anio','24',                                  'Meta anual de casos Surrogacy',                   true,  ''],
    // IMSS
    ['imss_uma_diaria',          '108.57',                              'Valor UMA diaria vigente (actualizar anual)',      true,  ''],
    ['imss_salario_minimo',      '278.80',                              'Salario mínimo diario vigente (actualizar)',       true,  ''],
    // Notificaciones
    ['notif_email_admin',        'enrique@hestiafertility.com',         'Email receptor de alertas del sistema',           true,  ''],
    ['notif_alertas_activas',    'true',                                'Activar envío de alertas automáticas',            true,  ''],
    ['notif_dias_alerta_cxp',    '7',                                   'Días de anticipación para alertar CxP vencidas',  true,  ''],
    // Sistema
    ['sistema_mantenimiento',    'false',                               'Modo mantenimiento (bloquea acceso)',              true,  ''],
    ['sistema_version_min',      '2026-01-01',                         'Versión mínima requerida del frontend',           true,  ''],
    ['sistema_max_tokens_activos','100',                                'Máximo de tokens activos simultáneos',            true,  '']
  ];
  sh.getRange(1, 1, 1, HDRS.length).setValues([HDRS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  if (DEFAULTS.length > 0) sh.getRange(2, 1, DEFAULTS.length, HDRS.length).setValues(DEFAULTS);
  sh.setColumnWidths(1, HDRS.length, [220, 280, 380, 60, 160]);
  try {
    var actRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox().setAllowInvalid(false).build();
    sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setDataValidation(actRule);
    var banding = sh.getRange(1,1,sh.getMaxRows(),HDRS.length);
    sh.getBandings().forEach(function(b){ b.remove(); });
    banding.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch(e) {}
  return sh;
}

/* ══════════════════════════════════════════════════════════════
   HELPERS PÚBLICOS — usables desde cualquier módulo del ERP
   ══════════════════════════════════════════════════════════════ */

/**
 * Lee un parámetro del sistema.
 * Primero busca en Script Properties (caché nocturna).
 * Si no existe, abre la hoja Configuracion como fallback.
 * @param {string} clave   — nombre del parámetro
 * @param {*}      default_  — valor si no se encuentra
 */
function getParam(clave, default_) {
  try {
    var val = PropertiesService.getScriptProperties().getProperty(CFG_PREFIX + clave);
    if (val !== null && val !== undefined && val !== '') return val;
  } catch(e) {}
  // Fallback: leer directo de la hoja (más lento, solo en caso de frio)
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var sh  = ss.getSheetByName('Configuracion');
    if (!sh) return default_ !== undefined ? default_ : '';
    var raw = sh.getDataRange().getValues();
    for (var r = 1; r < raw.length; r++) {
      if (String(raw[r][0]).trim() === clave) return String(raw[r][1]).trim();
    }
  } catch(e) {}
  return default_ !== undefined ? default_ : '';
}

/** Versión numérica de getParam */
function getParamNum(clave, default_) {
  var v = parseFloat(getParam(clave, ''));
  return isNaN(v) ? (default_ !== undefined ? default_ : 0) : v;
}

/** Versión booleana de getParam */
function getParamBool(clave, default_) {
  var v = getParam(clave, '').toLowerCase();
  if (v === 'true' || v === '1' || v === 'si' || v === 'sí') return true;
  if (v === 'false' || v === '0' || v === 'no') return false;
  return default_ !== undefined ? default_ : false;
}

/**
 * Verifica si un email es usuario activo del sistema.
 * Lee del caché de Script Properties (actualizado cada noche).
 */
function isUsuarioActivo(email) {
  try {
    var cached = PropertiesService.getScriptProperties().getProperty(CFG_PREFIX + '_usuarios');
    if (!cached) return null; // caché fría: no podemos responder
    var entries = cached.split('|');
    var emailL  = email.toLowerCase();
    for (var i = 0; i < entries.length; i++) {
      var parts = entries[i].split(':');
      if (parts[0] === emailL) return { email: parts[0], rol: parts[1] || 'viewer' };
    }
    return null;
  } catch(e) { return null; }
}

/** Limpia toda la caché de CacheService del script */
function clearSystemCache() {
  try {
    var cache = CacheService.getScriptCache();
    var known = ['erp_menu_v1','erp_menu_v2','erp_menu_v3','erp_banks_v1'];
    cache.removeAll(known);
    return { ok: true, message: 'Caché limpiada' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

/** Devuelve el estado actual del caché y de la última ejecución */
function estadoScheduler() {
  var props = PropertiesService.getScriptProperties();
  return {
    lastRun:     props.getProperty(CFG_LAST_RUN) || 'nunca',
    version:     props.getProperty(CFG_VERSION)  || 'sin versión',
    anio:        props.getProperty(CFG_PREFIX + 'anio_actual') || '',
    mes:         props.getProperty(CFG_PREFIX + 'mes_nombre')  || '',
    periodo:     props.getProperty(CFG_PREFIX + 'periodo_id')  || '',
    mantenimiento: props.getProperty(CFG_PREFIX + 'sistema_mantenimiento') || 'false',
    triggers: ScriptApp.getProjectTriggers()
      .filter(function(t){ return t.getHandlerFunction() === 'actualizarConfiguracionSistema'; })
      .map(function(t){ return { id: t.getUniqueId(), tipo: t.getEventType().toString() }; })
  };
}

/* ══════════════════════════════════════════════════════════════
   UTILITIES PRIVADAS
   ══════════════════════════════════════════════════════════════ */

/* ══════════════════════════════════════════════════════════════
   REGISTRO DE TAREAS PROGRAMADAS (extensible) — Panel de Control
   Para agregar una tarea nueva: añade una entrada en _schedTasks()
   con su handler (función real) y su tipo/hora default.
   ══════════════════════════════════════════════════════════════ */
function _schedTasks() {
  return [
    { id:'config_nocturna', handler:'actualizarConfiguracionSistema',
      nombre:'Actualización nocturna de configuración',
      desc:'Refresca caché de configuración, usuarios, períodos y valida conectividad de las hojas.',
      tipo:'diaria', horaDefault:23, minuto:0 },
    { id:'recordatorio_gastos_fijos', handler:'recordatorioGastosFijos',
      nombre:'Recordatorio de gastos fijos',
      desc:'El día 1 de cada mes envía un correo con los gastos fijos pendientes de programar.',
      tipo:'mensual', dia:1, horaDefault:8, minuto:0 },
    { id:'recordatorios_del_dia', handler:'recordatoriosDelDia',
      nombre:'Recordatorios del día',
      desc:'Cada mañana envía a cada usuario un correo con sus recordatorios (agenda) del día.',
      tipo:'diaria', horaDefault:7, minuto:30 },
    { id:'expediente_vigencias', handler:'expedienteChequeoNocturno',
      nombre:'Expediente médicos — avisos de vigencia',
      desc:'Cada mañana notifica a los médicos con documentos vencidos o por vencer (correo + recordatorio interno).',
      tipo:'diaria', horaDefault:8, minuto:15 }
  ];
}
function _schedGetCfg(id, key, def){
  var v = PropertiesService.getScriptProperties().getProperty('sched_'+id+'_'+key);
  return (v===null||v===undefined||v==='') ? def : v;
}
function _schedSetCfg(id, key, val){
  PropertiesService.getScriptProperties().setProperty('sched_'+id+'_'+key, String(val));
}

function readScheduledTasks() {
  try {
    var tasks = _schedTasks();
    // getProjectTriggers requiere el scope script.scriptapp. Si no está autorizado,
    // NO tronamos: devolvemos las tareas con estado desconocido y un flag.
    var triggers = [], scopeOk = true;
    try { triggers = ScriptApp.getProjectTriggers(); }
    catch(se) { scopeOk = false; }
    var lastRun = PropertiesService.getScriptProperties().getProperty(CFG_LAST_RUN) || '';
    var out = tasks.map(function(t){
      var hora   = parseInt(_schedGetCfg(t.id,'hora',t.horaDefault),10);
      var minuto = parseInt(_schedGetCfg(t.id,'minuto',t.minuto||0),10);
      var activo = String(_schedGetCfg(t.id,'activo','true')) !== 'false';
      var tipo   = _schedGetCfg(t.id,'tipo', t.tipo||'diaria');
      var dia    = parseInt(_schedGetCfg(t.id,'dia', t.dia||1),10); if (isNaN(dia)) dia = t.dia||1;
      var instalado = scopeOk ? triggers.some(function(tr){ return tr.getHandlerFunction()===t.handler; }) : null;
      return { id:t.id, nombre:t.nombre, desc:t.desc, tipo:tipo, handler:t.handler,
               dia:dia, hora:hora, minuto:minuto, activo:activo, instalado:instalado,
               lastRun: (t.handler==='actualizarConfiguracionSistema') ? lastRun : '' };
    });
    return { ok:true, tasks:out, scopeOk:scopeOk };
  } catch(ex){ return { ok:false, error:ex.message, tasks:[] }; }
}

function _schedWeekDay(n) {
  var W = [null, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY,
           ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.SATURDAY, ScriptApp.WeekDay.SUNDAY];
  return W[parseInt(n,10)] || ScriptApp.WeekDay.MONDAY;
}
// cfg = {tipo:'diaria'|'semanal'|'mensual', dia, hora, minuto, activo}
function _schedInstallTrigger(t, cfg) {
  ScriptApp.getProjectTriggers().forEach(function(tr){
    if (tr.getHandlerFunction()===t.handler) ScriptApp.deleteTrigger(tr);
  });
  if (!cfg.activo) return;
  var b = ScriptApp.newTrigger(t.handler).timeBased();
  if (cfg.tipo==='mensual') {
    b.onMonthDay(Math.max(1, Math.min(28, parseInt(cfg.dia,10)||1))).atHour(cfg.hora).create();
  } else if (cfg.tipo==='semanal') {
    b.onWeekDay(_schedWeekDay(cfg.dia)).atHour(cfg.hora).create();
  } else {
    b.everyDays(1).atHour(cfg.hora).nearMinute(cfg.minuto).create();
  }
}

function updateScheduledTask(body) {
  try {
    var t = _schedTasks().filter(function(x){ return x.id===body.id; })[0];
    if (!t) return { ok:false, error:'Tarea desconocida: '+body.id };
    var hora   = Math.max(0, Math.min(23, parseInt(body.hora,10)));
    if (isNaN(hora)) hora = t.horaDefault;
    var minuto = Math.max(0, Math.min(59, parseInt(body.minuto,10)));
    if (isNaN(minuto)) minuto = 0;
    var activo = !(body.activo===false || String(body.activo)==='false');
    var tipo = (body.tipo==='mensual'||body.tipo==='semanal'||body.tipo==='diaria') ? body.tipo : (t.tipo||'diaria');
    var dia = parseInt(body.dia,10); if (isNaN(dia)) dia = t.dia||1;
    // La configuración (tipo/día/hora/min/activo) se guarda siempre (no requiere permisos).
    _schedSetCfg(t.id,'hora',hora);
    _schedSetCfg(t.id,'minuto',minuto);
    _schedSetCfg(t.id,'tipo',tipo);
    _schedSetCfg(t.id,'dia',dia);
    _schedSetCfg(t.id,'activo',activo);
    // Instalar el disparador SÍ requiere el scope script.scriptapp; si falta, lo reportamos
    // sin tronar (la config queda guardada y se aplicará al autorizar/instalar).
    var cfg = {tipo:tipo, dia:dia, hora:hora, minuto:minuto, activo:activo};
    var scopeOk = true, instErr = '';
    try { _schedInstallTrigger(t, cfg); }
    catch(se) { scopeOk = false; instErr = se.message; }
    return { ok:true, id:t.id, tipo:tipo, dia:dia, hora:hora, minuto:minuto, activo:activo, scopeOk:scopeOk, instErr:instErr };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Instala (o reinstala) TODOS los triggers según su configuración.
   Correr UNA vez desde el editor; luego se ajusta desde el Panel de Control. */
function setupScheduledTriggers() {
  var tasks = _schedTasks(), res = [];
  tasks.forEach(function(t){
    var cfg = {
      tipo:   _schedGetCfg(t.id,'tipo', t.tipo||'diaria'),
      dia:    parseInt(_schedGetCfg(t.id,'dia', t.dia||1),10) || 1,
      hora:   parseInt(_schedGetCfg(t.id,'hora', t.horaDefault),10),
      minuto: parseInt(_schedGetCfg(t.id,'minuto', t.minuto||0),10),
      activo: String(_schedGetCfg(t.id,'activo','true')) !== 'false'
    };
    _schedInstallTrigger(t, cfg);
    res.push(t.id+' ('+cfg.tipo+') @ '+cfg.hora+':'+String(cfg.minuto).padStart(2,'0')+(cfg.activo?'':' (pausada)'));
  });
  Logger.log('[scheduler] Triggers instalados: '+res.join(', '));
  return 'Instalados: '+res.join(', ');
}

/* Handler mensual: avisa por correo de los gastos fijos por programar. */
function recordatorioGastosFijos() {
  try {
    if (typeof readGastosFijosPropuestas !== 'function') return;
    var prop = readGastosFijosPropuestas('');
    if (!prop || !prop.ok || !(prop.pendientes > 0)) return;
    var email = getParam('notif_email_admin', 'enrique@hestiafertility.com');
    var lista = prop.propuestas.map(function(p){
      return '• ' + p.proveedor + ' — ' + p.concepto + '  (vence ' + p.vencimiento + ')';
    }).join('\n');
    MailApp.sendEmail(email,
      'Gastos fijos por programar — ' + prop.periodo,
      'Tienes ' + prop.pendientes + ' gasto(s) fijo(s) por programar este mes:\n\n' + lista +
      '\n\nEntra al ERP → Gastos Fijos → Programación del mes para revisarlos y mandarlos a Cuentas por Pagar.');
  } catch(e) { Logger.log('[scheduler] recordatorioGastosFijos: ' + e.message); }
}

function _fmtDateLocal(d) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function _isoLocal(d) {
  return _fmtDateLocal(d) + ' ' + String(d.getHours()).padStart(2,'0') + ':' +
         String(d.getMinutes()).padStart(2,'0') + ':' + String(d.getSeconds()).padStart(2,'0');
}
