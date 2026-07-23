/* ==============================================================
   capture.gs — Módulo Captura & Datos Mensuales
   --------------------------------------------------------------
   Insert/Update genérico, Caja Chica, Mensual, Captura
   Proyecto Google Apps Script — Hestia Fertility ERP
   Todas las constantes vienen de config.gs (mismo proyecto)
   ============================================================== */

function readMensualData(ss, fechaInicio, fechaFin, viewId, sucursal) {
  sucursal = sucursal || 'Todas';
  var label = fechaInicio.slice(0,7) + ' → ' + fechaFin.slice(0,7);
  return {
    view:          viewId,
    periodo:       label,
    fechaInicio:   fechaInicio,
    fechaFin:      fechaFin,
    todos:         readMensual(ss, 'Mensual_Todos',         fechaInicio, fechaFin, sucursal),
    local:         readMensual(ss, 'Mensual_Local',         fechaInicio, fechaFin, sucursal),
    internacional: readMensual(ss, 'Mensual_Internacional', fechaInicio, fechaFin, sucursal),
    servicios:     readServicios(ss),
    funnel:        readFunnel(ss),
    alertas:       readAlertas(ss),
    donut:         readDonut(ss),
    cashflow:      readCashFlow(ss, fechaInicio, fechaFin, sucursal),
    costos:        readCostos(ss),
    paisesOrigen:  readPaisesOrigen(ss, fechaInicio, fechaFin, sucursal),
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

/* Detecta la fila de encabezado de una hoja de captura (igual lógica que readCapturaData)
   y devuelve { headers, dataStart } usando posiciones absolutas de columna (incluye Periodo si existe). */
function getSheetHeaders(sheet) {
  var allRows = sheet.getDataRange().getValues();
  function countFilled(row) {
    return row.filter(function(c) { return String(c).trim() !== ''; }).length;
  }
  var r0 = allRows.length > 0 ? countFilled(allRows[0]) : 0;
  var r1 = allRows.length > 1 ? countFilled(allRows[1]) : 0;
  var headerRow, dataStart;

  if (r0 === 0) {
    headerRow = allRows[1] || [];
    dataStart = 2;
  } else if (r0 > 0 && r1 > 0) {
    var complementario = allRows[0].every(function(v, i) {
      var v0 = String(v).trim();
      var v1 = String((allRows[1][i] !== undefined ? allRows[1][i] : '')).trim();
      return !(v0 && v1);
    });
    if (complementario) {
      headerRow = allRows[0].map(function(v, i) {
        return String(v).trim() || String(allRows[1][i] !== undefined ? allRows[1][i] : '').trim();
      });
      dataStart = 2;
    } else {
      headerRow = allRows[0];
      dataStart = 1;
    }
  } else {
    headerRow = allRows[0] || [];
    dataStart = 1;
  }

  var headers = headerRow.map(function(h) { return String(h).trim(); });
  return { headers: headers, dataStart: dataStart };
}

/* insert: ?action=insert&sheet=X&Campo1=valor1&... → agrega una fila nueva al final */
function insertRow(ss, e) {
  var sheetName = (e && e.parameter.sheet) || '';
  if (SHEET_ALIASES[sheetName]) sheetName = SHEET_ALIASES[sheetName];
  if (!sheetName) return { error: 'sheet es requerido' };
  var capturaId = getCapturaId(sheetName);
  // Lock: hace ATÓMICO "revisar duplicado + calcular folio + escribir" — evita
  // folios HEC duplicados y que dos altas simultáneas pisen la misma fila libre
  // (mismo patrón que nomina/gastosfijos/cobranza). Los inserts son poco
  // frecuentes, así que serializarlos no afecta el uso.
  var _insLock = LockService.getScriptLock();
  try { _insLock.waitLock(10000); } catch(eLk) { return { error: 'El sistema está ocupado, intenta de nuevo en unos segundos.' }; }
  try {
    var ssIns = SpreadsheetApp.openById(capturaId);
    var shIns = findSheet(ssIns, sheetName);
    if (!shIns) return { error: 'Hoja no encontrada: ' + sheetName };
    var hdrInfo = getSheetHeaders(shIns);
    var hdrs = hdrInfo.headers;

    // Validación de duplicados en Pacientes: ni el mismo nombre ni el mismo
    // correo. Si el correo ya existe, se dice QUÉ paciente lo tiene (para no
    // duplicar cuando dieron otro nombre con el mismo correo).
    if (sheetName.trim().toLowerCase() === 'pacientes') {
      var nombreIdx = -1, emailIdx = -1, idIdx = -1;
      for (var hi = 0; hi < hdrs.length; hi++) {
        var hl = hdrs[hi].toLowerCase();
        if (nombreIdx < 0 && hl.indexOf('nombre') > -1) nombreIdx = hi;
        if (emailIdx < 0 && (hl.indexOf('mail') > -1 || hl === 'correo')) emailIdx = hi;
        if (idIdx < 0 && hl === 'id') idIdx = hi;
      }
      var _fzP = String((e.parameter.forzar) || '').toLowerCase();
      var forzarPac = (_fzP === '1' || _fzP === 'true' || _fzP === 'si');
      var nombreRaw   = nombreIdx > -1 ? String(e.parameter[hdrs[nombreIdx]] || '').trim() : '';
      var emailNuevo  = emailIdx  > -1 ? String(e.parameter[hdrs[emailIdx]]  || '').trim().toLowerCase() : '';
      if (nombreRaw || emailNuevo) {
        var allData = shIns.getDataRange().getValues();
        // 1) Correo EXACTO → bloqueo fuerte (mismo correo = casi seguro la misma persona).
        if (emailNuevo) {
          for (var ri = hdrInfo.dataStart; ri < allData.length; ri++) {
            if (String(allData[ri][emailIdx] || '').trim().toLowerCase() === emailNuevo) {
              var quienId  = idIdx     > -1 ? String(allData[ri][idIdx] || '')     : '';
              var quienNom = nombreIdx > -1 ? String(allData[ri][nombreIdx] || '') : '';
              return {
                error: 'El correo "' + e.parameter[hdrs[emailIdx]] + '" ya está registrado' +
                       ((quienId || quienNom) ? ' en el paciente ' + (quienId ? quienId + ' — ' : '') + quienNom : '') + '.',
                duplicado: true, duplicadoCorreo: true, pacienteId: quienId, pacienteNombre: quienNom
              };
            }
          }
        }
        // 2) NOMBRE → candado DIFUSO (acentos / apellidos invertidos / espacios /
        //    erratas / iniciales "Melina PA"). Si hay candidatos y no se forzó, se
        //    pide confirmar — NO bloquea homónimos reales (el usuario puede forzar).
        if (!forzarPac && nombreRaw) {
          var cands = _pacDupCandidatos(nombreRaw, allData, hdrInfo.dataStart, nombreIdx, idIdx, emailIdx, null, 85);
          if (cands.length) {
            return { needConfirm: true, duplicado: true, candidatos: cands,
              error: 'Hay ' + cands.length + ' paciente(s) muy parecido(s) a "' + nombreRaw + '". Revisa si es la misma persona antes de crear otro.' };
          }
        }
      }
    }

    var newRow = hdrs.map(function(h) {
      return (h && e.parameter[h] !== undefined) ? e.parameter[h] : '';
    });
    // PACIENTES (bajo el lock): se (re)lee la hoja, se calcula el folio
    // AUTORITATIVO en el servidor — NO se confía en el ID del cliente, así dos
    // altas simultáneas nunca generan el mismo HEC — y se reutiliza una fila con
    // el NOMBRE en blanco (folio liberado por un borrado) en vez de dejar el hueco
    // al final. Seguro para ingresos: se ligan por NOMBRE, y una fila sin nombre
    // no tiene operaciones asociadas.
    if (sheetName.trim().toLowerCase() === 'pacientes' && typeof nombreIdx !== 'undefined' && nombreIdx > -1) {
      var allR = shIns.getDataRange().getValues();
      var maxN = 0, prefFolio = 'HEC';
      for (var mr = hdrInfo.dataStart; mr < allR.length; mr++) {
        var idc = (idIdx > -1) ? String(allR[mr][idIdx] || '').trim() : '';
        var mm = idc.match(/^([A-Za-z]+)?\D*?(\d+)\s*$/);
        if (mm && mm[2]) { var nn = parseInt(mm[2], 10); if (nn > maxN) maxN = nn; if (mm[1]) prefFolio = mm[1]; }
      }
      var folioNuevo = prefFolio + '-' + String(maxN + 1).padStart(3, '0');
      // Reutilizar la primera fila SIN nombre (folio liberado).
      for (var rr = hdrInfo.dataStart; rr < allR.length; rr++) {
        if (String(allR[rr][nombreIdx] || '').trim()) continue;
        if (idIdx > -1) {
          var idExist = String(allR[rr][idIdx] || '').trim();
          newRow[idIdx] = idExist || folioNuevo;   // conserva su folio; si no tiene, uno nuevo
        }
        shIns.getRange(rr + 1, 1, 1, newRow.length).setValues([newRow]);
        invalidateViewCache(sheetName);
        return { success: true, rowNum: rr + 1, reutilizado: true, id: (idIdx > -1 ? newRow[idIdx] : '') };
      }
      // Sin fila libre → append con folio autoritativo (ignora el ID del cliente).
      if (idIdx > -1) newRow[idIdx] = folioNuevo;
    }
    shIns.appendRow(newRow);
    var rowNum = shIns.getLastRow();
    invalidateViewCache(sheetName);
    return { success: true, rowNum: rowNum };
  } catch(ex) {
    return { error: ex.message };
  } finally {
    try { _insLock.releaseLock(); } catch(eR) {}
  }
}

/* Candidatos de posible duplicado por NOMBRE (difuso). Reusa el motor de
   identificar.gs (_identSim: acentos + apellidos invertidos + erratas) y el
   matcher de iniciales de analisis.gs (_ecMismoPaciente: "Melina PA" ↔ "Melina
   Pérez Álvarez"). Opera sobre los datos YA leídos de la hoja Pacientes. El
   correo se enmascara si el rol no puede ver datos sensibles (_privVer).
   Devuelve top-6 con score >= min. */
function _pacDupCandidatos(nombreRaw, allData, dataStart, nombreIdx, idIdx, emailIdx, excludeRow, min) {
  var out = [];
  if (!nombreRaw || nombreIdx < 0 || typeof _identSim !== 'function') return out;
  var minSc = (min == null) ? 85 : min;
  var _ver = (typeof _privVer !== 'function') || _privVer();
  for (var ri = dataStart; ri < allData.length; ri++) {
    if (excludeRow != null && ri === excludeRow) continue;
    var nomCat = String(allData[ri][nombreIdx] || '').trim();
    if (!nomCat) continue;
    var sc = _identSim(nombreRaw, nomCat);
    var ab = (typeof _ecMismoPaciente === 'function') && _ecMismoPaciente(nombreRaw, nomCat);
    if (ab && sc < 90) sc = 90;   // iniciales = match fuerte aunque el score textual sea bajo
    if (sc >= minSc) {
      out.push({
        id: (idIdx > -1 ? String(allData[ri][idIdx] || '') : ''),
        nombre: nomCat,
        email: (_ver && emailIdx > -1 ? String(allData[ri][emailIdx] || '') : ''),
        score: sc, abrev: !!ab
      });
    }
  }
  out.sort(function(a, b){ return b.score - a.score; });
  return out.slice(0, 6);
}

/* ══════════════════════════════════════════════════════════════
   CAJA CHICA — hoja independiente con saldo corrido (columna TOTAL
   es fórmula en Sheets: Total_n = Total_(n-1) - Salida_n + Entrada_n).
   La pestaña activa lleva como nombre el año en curso (ej. "2026");
   muchas filas debajo de la última captura ya tienen la fórmula de
   TOTAL copiada hacia abajo en blanco, esperando captura futura —
   por eso nunca usamos appendRow() aquí, sino que buscamos la
   primera fila con FECHA y CONCEPTO vacíos para no romper el orden.
   ══════════════════════════════════════════════════════════════ */
/* ═══════════════════════ CONFIG DE CAJAS (multi-caja) ═══════════════════════
 * Varias cajas de efectivo en el MISMO libro CAJA_CHICA_SS_ID, una PESTAÑA por caja.
 * La 'principal' es la histórica (hoja 'Caja Chica') — NO se toca. Las demás (ej. la de
 * recepción) tienen nombre EDITABLE. Config en Script Property CAJAS_CFG = [{id,label,tab}].
 * tab vacío en 'principal' = usa la resolución histórica (Caja Chica / año). */
var CAJAS_CFG_KEY = 'CAJAS_CFG';
function _cajasCfg() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(CAJAS_CFG_KEY);
    var arr = raw ? JSON.parse(raw) : null;
    if (!arr || !arr.length) arr = [{ id: 'principal', label: 'Caja Chica General', tab: '' }];
    if (!arr.some(function(c){ return c.id === 'principal'; }))
      arr.unshift({ id: 'principal', label: 'Caja Chica General', tab: '' });
    return arr;
  } catch (e) { return [{ id: 'principal', label: 'Caja Chica General', tab: '' }]; }
}
function _cajasSave(arr) { PropertiesService.getScriptProperties().setProperty(CAJAS_CFG_KEY, JSON.stringify(arr)); }
function _cajaCol(n){ var s=''; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=(n-m-1)/26; } return s; }

/* ¿Este actor puede acceder/mover en esta caja?
 *  · Sin email (llamada interna del sistema, ej. banco→caja) → SÍ.
 *  · Con ver_datos_sensibles (admin/finanzas) → SÍ a todas.
 *  · Caja con lista `usuarios` vacía → abierta (SÍ). Con lista → solo esos correos. */
function _cajaGate(cajaId, email, token) {
  if (!email) return true;
  try { if (token && typeof _tokenHasPermission === 'function' && _tokenHasPermission(token, 'ver_datos_sensibles')) return true; } catch (e) {}
  var caja = _cajasCfg().filter(function(c){ return c.id === cajaId; })[0];
  if (!caja) return false;
  var us = (caja.usuarios || []).map(function(x){ return String(x).toLowerCase().trim(); });
  if (!us.length) return true;
  return us.indexOf(String(email).toLowerCase().trim()) > -1;
}
/* Lista de cajas VISIBLES para el actor (filtradas por acceso). Si es admin
 * (ver_datos_sensibles) también devuelve `all` con las listas de usuarios para el panel
 * de accesos. */
function readCajas(email, token) {
  try {
    var all = _cajasCfg();
    var esAdmin = false;
    try { esAdmin = !!(token && _tokenHasPermission(token, 'ver_datos_sensibles')); } catch(e){}
    var vis = all.filter(function(c){ return _cajaGate(c.id, email, token); })
                 .map(function(c){ return { id:c.id, label:c.label, tab:c.tab }; });
    return { ok:true, cajas:vis, admin:esAdmin, all: esAdmin ? all : undefined };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
/* Saldos de TODAS las cajas (para el Flujo de Efectivo: Total Disponible = bancos + Σ cajas).
 * Devuelve el saldo de cada caja + el total. No filtra por acceso: es un AGREGADO para el
 * flujo (los montos por caja los ve quien ve el Flujo; el detalle de movimientos sí se gatea). */
function readCajasSaldos() {
  try {
    var out = [], total = 0;
    _cajasCfg().forEach(function(c){
      var saldo = 0;
      try { saldo = Number(readCajaChicaData(c.id).saldoFinal) || 0; } catch(e){}
      out.push({ id:c.id, label:c.label, saldo:saldo });
      total += saldo;
    });
    return { ok:true, cajas:out, total: Math.round(total*100)/100 };
  } catch(ex){ return { ok:false, error:ex.message, cajas:[], total:0 }; }
}

/* Asignar los correos con acceso a una caja (lista vacía = abierta). Gated. */
function asignarUsuariosCaja(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'ver_datos_sensibles'))
      return { ok:false, error:'Sin permiso para configurar accesos de caja.' };
    var arr = _cajasCfg(), id = String(body.id||'').trim();
    var us = Array.isArray(body.usuarios) ? body.usuarios.map(function(x){ return String(x).trim(); }).filter(Boolean) : [];
    var found = false;
    arr.forEach(function(c){ if(c.id===id){ c.usuarios = us; found=true; } });
    if (!found) return { ok:false, error:'No existe la caja '+id+'.' };
    _cajasSave(arr);
    try{ logAudit(body.usuario||'', 'CajaChica', 'Asignar accesos', id, 'usuarios', '', us.join(', ')||'(abierta)'); }catch(e){}
    return { ok:true, cajas:arr };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Renombrar una caja (nombre EDITABLE). Gated ver_datos_sensibles. */
function renombrarCaja(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'ver_datos_sensibles'))
      return { ok:false, error:'Sin permiso para configurar cajas.' };
    var arr = _cajasCfg(), id = String(body.id||'').trim(), label = String(body.label||'').trim();
    if (!label) return { ok:false, error:'El nombre no puede ir vacío.' };
    var found = false;
    arr.forEach(function(c){ if(c.id===id){ c.label=label; found=true; } });
    if (!found) return { ok:false, error:'No existe la caja '+id+'.' };
    _cajasSave(arr);
    try{ logAudit(body.usuario||'', 'CajaChica', 'Renombrar caja', id, 'label', '', label); }catch(e){}
    return { ok:true, cajas:arr };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Crear una caja NUEVA (ej. "Caja Extra" de recepción) con su pestaña. Gated. */
function crearCaja(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'ver_datos_sensibles'))
      return { ok:false, error:'Sin permiso para crear cajas.' };
    var arr = _cajasCfg(), label = String(body.label||'').trim() || 'Caja Extra';
    // id estable derivado; tab estable (no cambia al renombrar).
    var n = arr.length; var id = 'caja' + (n); var tab = 'Caja_' + id;
    while (arr.some(function(c){ return c.id===id || c.tab===tab; })) { n++; id='caja'+n; tab='Caja_'+id; }
    var ss = SpreadsheetApp.openById(CAJA_CHICA_SS_ID);
    var sh = ss.getSheetByName(tab) || ss.insertSheet(tab);
    sh.clear();
    sh.getRange(1,1,1,5).setValues([['FECHA','CONCEPTO','SALIDA','ENTRADA','TOTAL']]).setFontWeight('bold').setBackground('#f3f4f6');
    var saldoIni = parseFloat(body.saldoInicial)||0;
    sh.getRange(2,1,1,5).setValues([['REMANENTE INICIAL','', '', '', saldoIni]]);
    sh.setFrozenRows(1);
    arr.push({ id:id, label:label, tab:tab });
    _cajasSave(arr);
    try{ logAudit(body.usuario||'', 'CajaChica', 'Crear caja', id, 'label', '', label+' · saldo '+saldoIni); }catch(e){}
    return { ok:true, caja:{id:id,label:label,tab:tab}, cajas:arr };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Resuelve la HOJA de una caja. cajaId falsy/'principal' → histórico (Caja Chica/año). */
function getCajaChicaSheet(cajaId) {
  var ss = SpreadsheetApp.openById(CAJA_CHICA_SS_ID);
  cajaId = String(cajaId || 'principal');
  if (cajaId !== 'principal') {
    var caja = _cajasCfg().filter(function(c){ return c.id===cajaId; })[0];
    if (caja && caja.tab) { var s = ss.getSheetByName(caja.tab); if (s) return s; }
  }
  var sh = ss.getSheetByName('Caja Chica');
  if (sh) return sh;
  var anioActual = String(new Date().getFullYear());
  sh = ss.getSheetByName(anioActual);
  if (sh) return sh;
  return ss.getSheets()[0]; // pestaña más reciente como último recurso
}

function readCajaChicaData(cajaId) {
  try {
    var sh = getCajaChicaSheet(cajaId);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { movimientos: [], saldoInicial: 0, saldoFinal: 0 };
    var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iSalida   = headers.indexOf('SALIDA');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    function fmtFechaCC(v) {
      if (!v) return '';
      if (v instanceof Date) {
        var dd = String(v.getDate()).padStart(2,'0');
        var mm = String(v.getMonth()+1).padStart(2,'0');
        return dd + '/' + mm + '/' + v.getFullYear();
      }
      return String(v).trim();
    }
    var rows = [];
    var saldoInicial = 0;
    for (var r = 1; r < data.length; r++) {
      var row      = data[r];
      var fecha    = fmtFechaCC(row[iFecha]);
      var concepto = String(row[iConcepto] || '').trim();
      if (!fecha && !concepto) continue; // fila reservada (solo fórmula de TOTAL), aún sin capturar
      var salida   = Number(row[iSalida])  || 0;
      var entrada  = Number(row[iEntrada]) || 0;
      var total    = Number(row[iTotal])   || 0;
      var esRemanente = /^REMANENTE/i.test(fecha);
      if (esRemanente) saldoInicial = total;
      rows.push({
        _rowNum:    r + 1,
        fecha:      esRemanente ? '' : fecha,
        concepto:   esRemanente ? fecha : concepto,
        esRemanente: esRemanente,
        salida:     salida,
        entrada:    entrada,
        total:      total
      });
    }

    var saldoFinal = rows.length ? rows[rows.length - 1].total : saldoInicial;

    // Resumen de gasto por periodo (admite DD/MM/YYYY y MM/DD/YYYY mezclados en la hoja)
    function parseFechaMx(f) {
      var m = f.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (!m) return null;
      var a = parseInt(m[1], 10), b = parseInt(m[2], 10), y = parseInt(m[3], 10);
      var day = a, month = b;
      if (a > 12) { day = a; month = b; }
      else if (b > 12) { day = b; month = a; }
      return new Date(y, month - 1, day);
    }
    var hoy        = new Date();
    var inicioHoy  = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    var inicio7d   = new Date(inicioHoy.getTime() - 6 * 24 * 60 * 60 * 1000);
    var inicioMes  = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    var gastoHoy = 0, gastoSemana = 0, gastoMes = 0, ingresoMes = 0;
    rows.forEach(function(m) {
      if (m.esRemanente) return;
      // Movimientos INTERNOS (traspaso entre cajas, reverso de una venta, ajuste por
      // arqueo) SÍ mueven el saldo pero NO son gasto/ingreso operativo → no cuentan en
      // los KPIs de gasto (si no, un traspaso o un reverso de $50k inflaría "Gasto del Mes").
      if (/traspaso|reverso|ajuste por arqueo/i.test(String(m.concepto || ''))) return;
      var d = parseFechaMx(m.fecha);
      if (!d) return;
      if (d >= inicioMes) { gastoMes += m.salida; ingresoMes += m.entrada; }
      if (d >= inicio7d)  gastoSemana += m.salida;
      if (d >= inicioHoy) gastoHoy += m.salida;
    });

    return {
      saldoInicial: saldoInicial,
      saldoFinal:   saldoFinal,
      gastoHoy:     gastoHoy,
      gastoSemana:  gastoSemana,
      gastoMes:     gastoMes,
      ingresoMes:   ingresoMes,
      movimientos:  rows.slice().reverse(), // más reciente primero
      updated:      new Date().toISOString()
    };
  } catch(ex) {
    return { error: ex.message, movimientos: [], saldoInicial: 0, saldoFinal: 0 };
  }
}

/* Parser robusto de montos de caja: "$ 4,853.44", "$ (3,000.00)", "$ - ", "#REF!". */
function _ccNum(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).trim();
  if (!s || /#REF|#ERROR|#N\/A|#VALUE|#DIV/i.test(s)) return 0;
  var neg = /^\(.*\)$/.test(s);
  s = s.replace(/[$\s,()]/g, '');
  if (s === '-' || s === '') return 0;
  var num = parseFloat(s);
  if (isNaN(num)) return 0;
  return neg ? -Math.abs(num) : num;
}
/* Recalcula TODA la columna TOTAL como VALOR corrido desde el REMANENTE. Auto-repara
 * fórmulas rotas (#REF!) o vacías. Devuelve el saldo final. Las filas placeholder vacías
 * (sin FECHA ni CONCEPTO) quedan con TOTAL vacío. */
function _cajaRecomputeTotal(sh) {
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return 0;
  var h = data[0].map(function(x){ return String(x).trim().toUpperCase(); });
  var iF=h.indexOf('FECHA'), iC=h.indexOf('CONCEPTO'), iS=h.indexOf('SALIDA'), iE=h.indexOf('ENTRADA'), iT=h.indexOf('TOTAL');
  if (iT<0) return 0;
  var saldo=0, out=[];
  for (var r=1;r<data.length;r++){
    var fecha=String(data[r][iF]||'').trim(), concepto=String(data[r][iC]||'').trim();
    if (/^REMANENTE/i.test(fecha)) { saldo=_ccNum(data[r][iT]); out.push([saldo]); continue; }
    if (!fecha && !concepto) { out.push(['']); continue; }
    saldo = Math.round((saldo - _ccNum(data[r][iS]) + _ccNum(data[r][iE]))*100)/100;
    out.push([saldo]);
  }
  if (out.length) sh.getRange(2, iT+1, out.length, 1).setValues(out);
  return saldo;
}

/* Escribe un movimiento en una caja: reusa placeholder o AGREGA fila, y luego RECALCULA
 * toda la columna TOTAL como valor corrido (robusto a fórmulas rotas). Cualquier caja. */
function _cajaEscribirMov(sh, mov) {
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
  var iFecha=headers.indexOf('FECHA'), iConcepto=headers.indexOf('CONCEPTO'),
      iSalida=headers.indexOf('SALIDA'), iEntrada=headers.indexOf('ENTRADA');
  var targetRow = -1;
  for (var r=1;r<data.length;r++){ if(!String(data[r][iFecha]||'').trim() && !String(data[r][iConcepto]||'').trim()){ targetRow=r+1; break; } }
  if (targetRow===-1) targetRow = sh.getLastRow()+1;
  sh.getRange(targetRow, iFecha+1).setValue(mov.fecha);
  sh.getRange(targetRow, iConcepto+1).setValue(mov.concepto);
  sh.getRange(targetRow, iSalida+1).setValue(mov.salida||0);
  sh.getRange(targetRow, iEntrada+1).setValue(mov.entrada||0);
  SpreadsheetApp.flush();
  var saldoFinal = _cajaRecomputeTotal(sh);
  return { targetRow: targetRow, saldoFinal: saldoFinal };
}

/* Reparar saldos de una caja (recalcula la columna TOTAL). Gated ver_datos_sensibles. */
function repararCajaChica(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'ver_datos_sensibles'))
      return { ok:false, error:'Sin permiso para reparar la caja.' };
    var sh = getCajaChicaSheet(body.caja || 'principal');
    var saldo = _cajaRecomputeTotal(sh);
    try{ logAudit(body.usuario||'', 'CajaChica', 'Reparar saldos', body.caja||'principal', 'saldoFinal', '', String(saldo)); }catch(e){}
    return { ok:true, saldoFinal:saldo };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

function insertCajaChicaRow(e) {
  try {
    var cajaId = (e && e.parameter && e.parameter.caja) || 'principal';
    var _tk = (e && e.parameter && e.parameter.token) || '';
    var _em = (typeof verifyToken === 'function') ? verifyToken(_tk) : '';
    if (!_cajaGate(cajaId, _em, _tk)) return { error: 'No tienes acceso a esta caja.' };
    var sh = getCajaChicaSheet(cajaId);
    var concepto = String(e.parameter['CONCEPTO'] || '').trim();
    var fecha    = String(e.parameter['FECHA']    || '').trim();
    var salida   = parseFloat(e.parameter['SALIDA'])  || 0;
    var entrada  = parseFloat(e.parameter['ENTRADA']) || 0;
    if (!concepto)            return { error: 'El concepto es requerido.' };
    if (!fecha)               return { error: 'La fecha es requerida.' };
    if (!salida && !entrada)  return { error: 'Captura un monto de salida o entrada.' };
    var res = _cajaEscribirMov(sh, { fecha:fecha, concepto:concepto, salida:salida, entrada:entrada });
    return { success: true, rowNum: res.targetRow, saldoFinal: res.saldoFinal };
  } catch(ex) {
    return { error: ex.message };
  }
}

function updateCajaChicaRow(body) {
  try {
    if (!_cajaGate(body.caja || 'principal', body.usuario, body.token)) return { error:'No tienes acceso a esta caja.' };
    var sh      = getCajaChicaSheet(body.caja || 'principal');
    var rowNum  = parseInt(body.rowNum);
    if (!rowNum || rowNum < 2) return { error: 'Número de fila inválido.' };
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iFecha    = headers.indexOf('FECHA');
    var iConcepto = headers.indexOf('CONCEPTO');
    var iSalida   = headers.indexOf('SALIDA');
    var iEntrada  = headers.indexOf('ENTRADA');
    var iTotal    = headers.indexOf('TOTAL');

    if (body.fecha)    sh.getRange(rowNum, iFecha + 1).setValue(body.fecha);
    if (body.concepto) sh.getRange(rowNum, iConcepto + 1).setValue(body.concepto);
    sh.getRange(rowNum, iSalida + 1).setValue(parseFloat(body.salida) || 0);
    sh.getRange(rowNum, iEntrada + 1).setValue(parseFloat(body.entrada) || 0);
    SpreadsheetApp.flush();
    var nuevoTotal = _cajaRecomputeTotal(sh);   // recalcula toda la columna (robusto a fórmulas rotas)
    return { ok: true, rowNum: rowNum, saldoFinal: Number(nuevoTotal) || 0 };
  } catch(ex) {
    return { error: ex.message };
  }
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

function readCapturaData(ss, nombreHoja, viewId, fechaInicio, fechaFin, sucursal) {
  sucursal = sucursal || 'Todas';
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

  // Filtro por sucursal — solo si la columna existe y el filtro está activo
  if (sucursal && sucursal !== 'Todas') {
    var sucHdrIdx = headers.map(function(h){ return h.toLowerCase(); }).indexOf('sucursal');
    if (sucHdrIdx >= 0) {
      rows = rows.filter(function(row) {
        var val = String(row['Sucursal'] || row['sucursal'] || '').trim();
        return val === '' || val === sucursal; // vacío = hereda todas las sucursales
      });
    }
  }

  return { view: viewId, fuente: nombreHoja, headers: headers, rows: rows,
           updated: new Date().toISOString() };
}

/* ══════════════════════════════════════════════════════════════
   LECTORES INDIVIDUALES
   ══════════════════════════════════════════════════════════════ */
/* Columnas: A=Sucursal | B=Fecha(YYYY-MM-DD) | C=Mes | D=Ingresos | E=Gastos | F=Ciclos | G=CAC | H=Margen */
function readMensual(ss, sheetName, fechaInicio, fechaFin, sucursal) {
  var hoja = ss.getSheetByName(sheetName);
  if (!hoja) return { meses:[], ingresos:[], gastos:[], ciclos:[], cac:[], margen:[] };
  var rows = hoja.getDataRange().getValues();
  if (rows.length < 2) return { meses:[], ingresos:[], gastos:[], ciclos:[], cac:[], margen:[] };
  // Detectar columnas por encabezado (soporta columnas en cualquier orden)
  var hdrs = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var iSuc    = hdrs.indexOf('sucursal');
  var iFecha  = hdrs.indexOf('fecha');  if (iFecha  < 0) iFecha  = 1;
  var iMes    = hdrs.indexOf('mes');    if (iMes    < 0) iMes    = 2;
  var iIngr   = hdrs.indexOf('ingresos'); if (iIngr < 0) iIngr   = 3;
  var iGast   = hdrs.indexOf('gastos');   if (iGast < 0) iGast   = 4;
  var iCiclos = hdrs.indexOf('ciclos');   if (iCiclos < 0) iCiclos = 5;
  var iCac    = hdrs.indexOf('cac');      if (iCac  < 0) iCac    = 6;
  var iMargen = hdrs.indexOf('margen');   if (iMargen < 0) iMargen = 7;
  var data = rows.slice(1).filter(function(r) {
    var f = String(r[iFecha]).trim();
    if (f < fechaInicio || f > fechaFin) return false;
    if (sucursal && sucursal !== 'Todas' && iSuc >= 0) {
      var s = String(r[iSuc] || '').trim();
      if (s && s !== sucursal) return false;
    }
    return true;
  });
  data.sort(function(a, b) { return String(a[iFecha]) < String(b[iFecha]) ? -1 : 1; });
  return {
    meses:    data.map(function(r) { return String(r[iMes]); }),
    ingresos: data.map(function(r) { return Number(r[iIngr]); }),
    gastos:   data.map(function(r) { return Number(r[iGast]); }),
    ciclos:   data.map(function(r) { return Number(r[iCiclos]); }),
    cac:      data.map(function(r) { return Number(r[iCac]); }),
    margen:   data.map(function(r) { return Number(r[iMargen]); })
  };
}

/* ── Funciones financieras definidas en finance.gs (mismo proyecto GAS) ──
   readCashFlow, readServicios, readFunnel, readAlertas, readDonut,
   readPaisesOrigen, readCostos, readEstadoResultados, readOperatingPL,
   _buildPLReport, readBanksData, saveBankRow, doPost
   ────────────────────────────────────────────────────────────────────────── */
function saveCajaChicaIngreso(body) {
  try {
    if (!_cajaGate(body.caja || 'principal', body.usuario, body.token)) return { ok:false, error:'No tienes acceso a esta caja.' };
    var sh = getCajaChicaSheet(body.caja || 'principal');
    var fecha    = String(body.fecha    || '').trim();
    var concepto = String(body.concepto || '').trim();
    var entrada  = parseFloat(body.entrada) || 0;
    var salida   = parseFloat(body.salida) || 0;   // permite también salida (traspaso/egreso)
    if (!fecha || !concepto || (!entrada && !salida)) {
      return { ok: false, error: 'fecha, concepto y monto son requeridos' };
    }
    var res = _cajaEscribirMov(sh, { fecha:fecha, concepto:concepto, salida:salida, entrada:entrada });
    return { ok: true, rowNum: res.targetRow, saldoFinal: res.saldoFinal };
  } catch(ex) {
    return { ok: false, error: ex.message };
  }
}

/* ─────────── CORTE / ARQUEO por caja + TRASPASO entre cajas ─────────── */
function _cortesSheet() {
  var ss = SpreadsheetApp.openById(CAJA_CHICA_SS_ID);
  var sh = ss.getSheetByName('Cortes_Caja');
  if (!sh) {
    sh = ss.insertSheet('Cortes_Caja');
    sh.appendRow(['Fecha','Caja','Usuario','SaldoEsperado','EfectivoContado','Diferencia','Observaciones','Timestamp']);
    sh.setFrozenRows(1); sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#f3f4f6');
  }
  return sh;
}
/* Corte/arqueo: compara el saldo esperado (libro) vs el efectivo contado. Si difiere y
 * ajustar=true, siembra un movimiento "Ajuste por arqueo" para que libros = realidad. */
function saveCorteCaja(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'ver_datos_sensibles'))
      return { ok:false, error:'Sin permiso para hacer cortes de caja.' };
    var cajaId = String(body.caja||'principal');
    var caja = _cajasCfg().filter(function(c){ return c.id===cajaId; })[0] || {id:cajaId,label:cajaId};
    var esperado = Number(readCajaChicaData(cajaId).saldoFinal)||0;
    var contado = parseFloat(body.contado);
    if (isNaN(contado)) return { ok:false, error:'Captura el efectivo contado.' };
    var dif = Math.round((contado - esperado)*100)/100;
    var fecha = body.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    _cortesSheet().appendRow([fecha, caja.label, body.usuario||'', esperado, contado, dif, String(body.observaciones||''), new Date()]);
    var ajustado = false;
    if (dif !== 0 && (body.ajustar === true || body.ajustar === 'true')) {
      var sh = getCajaChicaSheet(cajaId);
      _cajaEscribirMov(sh, { fecha:fecha, concepto:'Ajuste por arqueo ('+(dif>0?'sobrante':'faltante')+')',
        salida: dif<0 ? Math.abs(dif) : 0, entrada: dif>0 ? dif : 0 });
      ajustado = true;
    }
    try{ logAudit(body.usuario||'', 'CajaChica', 'Corte/arqueo', caja.label, 'diferencia', String(esperado), String(contado)+' (dif '+dif+(ajustado?', ajustado':'')+')'); }catch(e){}
    return { ok:true, esperado:esperado, contado:contado, diferencia:dif, ajustado:ajustado };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
function readCortesCaja(cajaId) {
  try {
    var sh = _cortesSheet(); var raw = sh.getDataRange().getValues();
    var caja = _cajasCfg().filter(function(c){ return c.id===cajaId; })[0];
    var lbl = caja ? caja.label : '';
    var rows = [];
    for (var i=1;i<raw.length;i++){
      if (cajaId && lbl && String(raw[i][1]).trim()!==lbl) continue;
      rows.push({ fecha:String(raw[i][0]), caja:String(raw[i][1]), usuario:String(raw[i][2]),
        esperado:Number(raw[i][3])||0, contado:Number(raw[i][4])||0, diferencia:Number(raw[i][5])||0,
        observaciones:String(raw[i][6]||'') });
    }
    return { ok:true, cortes:rows.reverse() };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
/* Traspaso entre cajas: SALIDA en origen + ENTRADA en destino, ligadas por referencia. */
function traspasoCaja(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'ver_datos_sensibles'))
      return { ok:false, error:'Sin permiso para traspasos de caja.' };
    var origen=String(body.origen||''), destino=String(body.destino||''), monto=parseFloat(body.monto)||0;
    if (!origen || !destino || origen===destino) return { ok:false, error:'Elige caja origen y destino distintas.' };
    if (monto<=0) return { ok:false, error:'Monto inválido.' };
    var cfg=_cajasCfg();
    var lo=(cfg.filter(function(c){return c.id===origen;})[0]||{}).label||origen;
    var ld=(cfg.filter(function(c){return c.id===destino;})[0]||{}).label||destino;
    var saldoO = Number(readCajaChicaData(origen).saldoFinal)||0;
    if (monto > saldoO + 0.001) return { ok:false, error:'La caja "'+lo+'" solo tiene '+saldoO.toFixed(2)+'.' };
    var fecha = body.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var ref = 'TRASPASO ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    _cajaEscribirMov(getCajaChicaSheet(origen),  { fecha:fecha, concepto:'Traspaso → '+ld+' ['+ref+']', salida:monto, entrada:0 });
    _cajaEscribirMov(getCajaChicaSheet(destino), { fecha:fecha, concepto:'Traspaso ← '+lo+' ['+ref+']', salida:0, entrada:monto });
    try{ logAudit(body.usuario||'', 'CajaChica', 'Traspaso', ref, lo+'→'+ld, '', String(monto)); }catch(e){}
    return { ok:true, ref:ref, origen:lo, destino:ld, monto:monto };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
