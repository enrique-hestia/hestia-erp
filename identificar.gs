/* ══════════════════════════════════════════════════════════════════════════
   IDENTIFICAR PACIENTES — ponerle nombre a los cobros "No localizado".
   ──────────────────────────────────────────────────────────────────────────
   POR QUÉ EXISTE
   La página de citas cobra por Mercado Pago pero NO captura el nombre de la
   paciente. Esos cobros entran a BD_Ingresos con el paciente literal
   "No localizado". Las únicas que saben de quién es cada pago son las
   vendedoras/enfermeras que lo cobraron — y ellas NO tienen (ni deben tener)
   permiso para entrar a Ingresos, porque ahí verían TODO el universo de
   operaciones.

   QUÉ HACE (y qué NO)
   Este módulo hace UNA sola cosa: cambiar el nombre de un "No localizado" por
   el nombre real. NO toca productos, montos, comisiones, CxCobrar ni la
   conciliación bancaria — no se mueve nada de eso. Solo el nombre.

   LOS DOS CANDADOS QUE LO MANTIENEN HONESTO
   1) Permiso 'identificar_pacientes' en los 4 endpoints.
   2) EL CANDADO CRÍTICO: asignarPacienteNoLocalizado() RELEE la fila y exige
      que el paciente SIGA siendo "No localizado". Sin esto, el permiso se
      convertiría en "editar el paciente de CUALQUIER ingreso". Con esto, solo
      puede escribir sobre filas que nadie ha identificado todavía.

   SUPERFICIE DE DATOS
   listarNoLocalizados() devuelve a propósito el mínimo indispensable para
   ubicar el cobro: OP, fecha, monto y método de pago. NADA más — ni productos,
   ni conceptos, ni otros pacientes, ni márgenes. Esta es la superficie COMPLETA
   que ve alguien sin permiso de ingresos: mantenerla mínima es el diseño.

   DEPENDENCIAS (scope global de Apps Script)
   finance.gs → _tokenHasPermission, logAudit, _maxIdNum, listaPacientesAll,
                _bankSheetByKey, _bankObsIdx, BD_INGRESOS_TAB, _ingIdDeAnio,
                PACIENTES_SS_ID, PAC_COL_LISTA
   facturacion.gs → _pacColIdx
   ══════════════════════════════════════════════════════════════════════════ */

/* Variantes que cuentan como "sin nombre". Se comparan YA normalizadas
   (minúsculas, sin acentos, sin puntuación y SIN espacios), así que
   'No Localizado', 'no localizado' y 'nolocalizado' caen todas en 'nolocalizado'.
   ▸ Para aceptar una variante nueva: agrégala aquí y ya. Nada más que tocar. */
var IDENT_NO_LOCALIZADO = [
  'nolocalizado',
  'sinnombre',
  'noidentificado',
  'sinidentificar',
  'pacientenolocalizado'
];

/* Umbral de duplicado: >= a esto, el alta pide confirmación explícita. */
var IDENT_SIM_UMBRAL = 90;

/* Columna del paciente en BD_Ingresos (1-indexed): D = Paciente.
   Ver BD_INGRESOS_HEADERS en finance.gs — OP(0),Linea(1),Fecha(2),Paciente(3). */
var IDENT_PAC_COL = 4;

/* ── Normalización ───────────────────────────────────────────────────────
   Sin .normalize('NFD') a propósito: el harness de validación corre en
   JScript (ES3) y no lo tiene. El reemplazo explícito funciona en ambos. */
function _identNorm(s) {
  var t = String(s == null ? '' : s).toLowerCase();
  t = t.replace(/[áàäâãå]/g, 'a').replace(/[éèëê]/g, 'e').replace(/[íìïî]/g, 'i');
  t = t.replace(/[óòöôõ]/g, 'o').replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n').replace(/ç/g, 'c');
  t = t.replace(/[^a-z0-9\s]/g, ' ');          // fuera puntuación
  t = t.replace(/\s+/g, ' ');
  t = t.replace(/^\s+|\s+$/g, '');             // trim (ES3 no tiene .trim)
  return t;
}

/* ¿Este texto es un "sin nombre"? Normaliza y colapsa espacios para que
   'No  Localizado' y 'nolocalizado' den lo mismo. */
function _identEsNoLocalizado(s) {
  var t = _identNorm(s).replace(/\s/g, '');
  if (!t) return true;                          // celda vacía = tampoco tiene nombre
  for (var i = 0; i < IDENT_NO_LOCALIZADO.length; i++) {
    if (t === IDENT_NO_LOCALIZADO[i]) return true;
  }
  return false;
}

/* Tokens ordenados: "López García, María" y "Maria Garcia Lopez" convergen.
   Así el orden de los apellidos deja de generar falsos negativos. */
function _identTokensSorted(s) {
  var t = _identNorm(s);
  if (!t) return '';
  var parts = t.split(' ');
  parts.sort();
  return parts.join(' ');
}

function _identBigrams(s) {
  var m = {};
  for (var i = 0; i < s.length - 1; i++) {
    var g = s.substr(i, 2);
    m[g] = (m[g] || 0) + 1;
  }
  return m;
}

/* Dice sobre bigramas de la cadena completa (tokens ordenados).
   Bueno para apellidos permutados y para "le sobra un apellido". */
function _identDice(a, b) {
  var x = _identTokensSorted(a), y = _identTokensSorted(b);
  if (!x || !y) return 0;
  if (x === y) return 100;                      // idénticos tras normalizar
  if (x.length < 2 || y.length < 2) return 0;   // sin bigramas que comparar
  var bx = _identBigrams(x), by = _identBigrams(y);
  var inter = 0, tx = 0, ty = 0, k;
  for (k in bx) { if (bx.hasOwnProperty(k)) tx += bx[k]; }
  for (k in by) {
    if (!by.hasOwnProperty(k)) continue;
    ty += by[k];
    if (bx[k]) inter += Math.min(bx[k], by[k]);
  }
  if (tx + ty === 0) return 0;
  return Math.round((2 * inter / (tx + ty)) * 100);
}

/* Distancia de Levenshtein (ES5 puro, dos filas — O(n) en memoria). */
function _identLev(a, b) {
  if (a === b) return 0;
  var la = a.length, lb = b.length;
  if (!la) return lb;
  if (!lb) return la;
  var prev = [], cur = [], i, j;
  for (j = 0; j <= lb; j++) prev[j] = j;
  for (i = 1; i <= la; i++) {
    cur[0] = i;
    for (j = 1; j <= lb; j++) {
      var cost = (a.charAt(i - 1) === b.charAt(j - 1)) ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, Math.min(cur[j - 1] + 1, prev[j - 1] + cost));
    }
    for (j = 0; j <= lb; j++) prev[j] = cur[j];
  }
  return prev[lb];
}

/* Levenshtein normalizado a 0-100 entre dos palabras sueltas. */
function _identLevSim(a, b) {
  var mx = Math.max(a.length, b.length);
  if (!mx) return 100;
  var s = Math.round((1 - _identLev(a, b) / mx) * 100);
  return s < 0 ? 0 : s;
}

/* Similitud palabra-por-palabra: empareja cada token con su mejor pareja del
   otro nombre y promedia, pesando por longitud (un apellido pesa más que un
   "de"). Los tokens que se quedan sin pareja cuentan como 0.
   ▸ POR QUÉ EXISTE, además de Dice: Dice sobre la cadena completa castiga
     demasiado un cambio de UNA letra en un apellido corto —
     "Ana Ruiz Perez" vs "Ana Ruiz Perea" daba 85 y se colaba por debajo del
     umbral de 90, dando de alta un duplicado en silencio. Comparando token a
     token da 92 y sí pide confirmación, que es justo el caso que preocupa al
     usuario ("casi los mismos apellidos"). */
function _identTokenSim(a, b) {
  var ta = _identNorm(a).split(' '), tb = _identNorm(b).split(' ');
  var la = [], lb = [], i, j;
  for (i = 0; i < ta.length; i++) if (ta[i]) la.push(ta[i]);
  for (i = 0; i < tb.length; i++) if (tb[i]) lb.push(tb[i]);
  if (!la.length || !lb.length) return 0;

  var corto = (la.length <= lb.length) ? la : lb;
  var largo = (la.length <= lb.length) ? lb : la;
  var usado = [], num = 0, den = 0;

  for (i = 0; i < corto.length; i++) {
    var best = -1, bestSc = -1;
    for (j = 0; j < largo.length; j++) {
      if (usado[j]) continue;
      var sc = _identLevSim(corto[i], largo[j]);
      if (sc > bestSc) { bestSc = sc; best = j; }
    }
    if (best === -1) { den += corto[i].length; continue; }
    usado[best] = true;
    var w = Math.max(corto[i].length, largo[best].length);
    num += bestSc * w;
    den += w;
  }
  for (j = 0; j < largo.length; j++) if (!usado[j]) den += largo[j].length;  // sobrantes = 0
  if (!den) return 0;
  return Math.round(num / den);
}

/* Similitud 0-100 final: lo MEJOR de los dos ángulos.
   Cada métrica cubre el punto ciego de la otra — Dice ve el nombre completo
   (apellidos permutados, apellido de más), los tokens ven la palabra exacta
   (una letra cambiada). Tomar el máximo hace que un duplicado tenga que
   engañar a las dos para colarse, y equivocarse hacia "pregunta de más"
   cuesta un clic; equivocarse hacia "alta en silencio" cuesta un duplicado. */
function _identSim(a, b) {
  var d = _identDice(a, b);
  var t = _identTokenSim(a, b);
  return d > t ? d : t;
}

/* Candidatos del catálogo por encima de minScore, ordenados desc.
   Interno: sin token. Los endpoints públicos son los que exigen permiso. */
function _identSimilares(nombre, minScore) {
  var min = (minScore === undefined || minScore === null) ? 60 : minScore;
  var lp = (typeof listaPacientesAll === 'function') ? listaPacientesAll() : null;
  if (!lp || !lp.ok || !lp.pacientes) return [];
  var out = [];
  for (var i = 0; i < lp.pacientes.length; i++) {
    var p = lp.pacientes[i];
    var sc = _identSim(nombre, p.nombre);
    if (sc >= min) out.push({ id: p.id, nombre: p.nombre, email: p.email || '', score: sc });
  }
  out.sort(function (a, b) { return b.score - a.score; });
  return out.slice(0, 8);
}

/* ── Hoja BD_Ingresos del año pedido ─────────────────────────────────── */
function _identSheetIngresos(anio) {
  var ss = SpreadsheetApp.openById(_ingIdDeAnio(anio));
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === BD_INGRESOS_TAB) return sheets[i];
  }
  return null;
}

function _identFecha(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var mm = String(v.getMonth() + 1), dd = String(v.getDate());
    if (mm.length < 2) mm = '0' + mm;
    if (dd.length < 2) dd = '0' + dd;
    return v.getFullYear() + '-' + mm + '-' + dd;
  }
  return String(v).substring(0, 10);
}

function _identNum(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v == null ? '' : v).replace(/[$,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

/* ══════════════════════════════════════════════════════════════════════
   1) LISTAR — los cobros que esperan nombre.
   Devuelve SOLO: op, fecha, monto (total de la operación) y método de pago.
   Deliberadamente nada más: esta respuesta es todo lo que verá alguien que
   no tiene permiso de Ingresos.
   Filtros: anio, fechaIni, fechaFin (YYYY-MM-DD), montoMin, montoMax.
   ══════════════════════════════════════════════════════════════════════ */
function listarNoLocalizados(params) {
  try {
    params = params || {};
    if (!_tokenHasPermission(params.token || '', 'identificar_pacientes'))
      return { ok: false, error: 'Sin autorización (identificar_pacientes).' };

    var anio = parseInt(params.anio, 10) || new Date().getFullYear();
    var sh = _identSheetIngresos(anio);
    if (!sh) return { ok: false, error: 'BD_Ingresos no encontrada para ' + anio + '.' };

    var data = sh.getDataRange().getValues();
    var fIni = String(params.fechaIni || '').substring(0, 10);
    var fFin = String(params.fechaFin || '').substring(0, 10);
    var mMin = (params.montoMin === '' || params.montoMin === undefined || params.montoMin === null)
      ? null : parseFloat(params.montoMin);
    var mMax = (params.montoMax === '' || params.montoMax === undefined || params.montoMax === null)
      ? null : parseFloat(params.montoMax);
    if (mMin !== null && isNaN(mMin)) mMin = null;
    if (mMax !== null && isNaN(mMax)) mMax = null;

    // Agrupar por OP: una operación son varias líneas; la paciente ve UNA fila.
    var ops = {}, orden = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var op = String(r[0] || '').replace(/^\s+|\s+$/g, '');
      if (!op) continue;
      if (!_identEsNoLocalizado(r[IDENT_PAC_COL - 1])) continue;

      if (!ops[op]) {
        ops[op] = { op: op, fecha: _identFecha(r[2]), monto: 0, metodoPago: '' };
        orden.push(op);
      }
      ops[op].monto += _identNum(r[9]);                       // TotalPagar
      var fp = String(r[12] || '').replace(/^\s+|\s+$/g, ''); // FormaPago
      if (fp && ops[op].metodoPago.indexOf(fp) === -1)
        ops[op].metodoPago = ops[op].metodoPago ? (ops[op].metodoPago + ' + ' + fp) : fp;
    }

    var out = [];
    for (var k = 0; k < orden.length; k++) {
      var o = ops[orden[k]];
      if (fIni && o.fecha && o.fecha < fIni) continue;
      if (fFin && o.fecha && o.fecha > fFin) continue;
      if (mMin !== null && o.monto < mMin) continue;
      if (mMax !== null && o.monto > mMax) continue;
      out.push({ op: o.op, fecha: o.fecha, monto: Math.round(o.monto * 100) / 100, metodoPago: o.metodoPago });
    }
    // Más recientes primero: es donde ella va a reconocer a su paciente.
    out.sort(function (a, b) { return a.fecha < b.fecha ? 1 : (a.fecha > b.fecha ? -1 : 0); });

    return { ok: true, anio: String(anio), pendientes: out, total: out.length };
  } catch (ex) {
    return { ok: false, error: ex.message, pendientes: [] };
  }
}

/* ══════════════════════════════════════════════════════════════════════
   2) ASIGNAR — el corazón. Escribe SOLO la celda del paciente.
   NO llama a updateIngreso / updateIngresoConBancos: esas BORRAN y RE-CREAN
   las filas de banco. Aquí eso está prohibido — no se mueve dinero, solo el
   nombre.
   ══════════════════════════════════════════════════════════════════════ */
function asignarPacienteNoLocalizado(body) {
  body = body || {};
  if (!_tokenHasPermission(body.token || '', 'identificar_pacientes'))
    return { ok: false, error: 'Sin autorización (identificar_pacientes).' };

  var opId = String(body.op || '').replace(/^\s+|\s+$/g, '');
  var nuevo = String(body.paciente || '').replace(/^\s+|\s+$/g, '');
  if (!opId) return { ok: false, error: 'Falta la operación.' };
  if (!nuevo) return { ok: false, error: 'Falta el nombre del paciente.' };
  if (_identEsNoLocalizado(nuevo))
    return { ok: false, error: 'Ese no es un nombre real. Escribe el nombre de la paciente.' };

  var usuario = body.usuario || '';
  var anio = parseInt(body.anio, 10) || new Date().getFullYear();

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (eLk) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo en unos segundos.' }; }

  try {
    var sh = _identSheetIngresos(anio);
    if (!sh) return { ok: false, error: 'BD_Ingresos no encontrada para ' + anio + '.' };

    var data = sh.getDataRange().getValues();
    var filas = [], viejo = '';
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').replace(/^\s+|\s+$/g, '') !== opId) continue;
      filas.push(i + 1);                                   // fila real (1-indexed)
      if (!viejo) viejo = String(data[i][IDENT_PAC_COL - 1] || '');
    }
    if (!filas.length) return { ok: false, error: 'No se encontró la operación ' + opId + '.' };

    /* ─────────────────────────────────────────────────────────────────
       EL CANDADO. Se relee la fila AQUÍ, dentro del lock, y se exige que
       el paciente SIGA siendo "No localizado". Dos motivos:
       · Carrera: otra persona pudo identificarla mientras esta lista
         estaba abierta en pantalla.
       · Seguridad: sin esto, 'identificar_pacientes' se convertiría en
         "renombrar al paciente de CUALQUIER ingreso". Este es el único
         candado que impide esa escalada de privilegio. NO QUITAR.
       ───────────────────────────────────────────────────────────────── */
    for (var f = 0; f < filas.length; f++) {
      var actual = data[filas[f] - 1][IDENT_PAC_COL - 1];
      if (!_identEsNoLocalizado(actual)) {
        return {
          ok: false, yaIdentificada: true,
          error: 'Esa operación ya fue identificada como "' +
                 String(actual).replace(/^\s+|\s+$/g, '') + '". Solo se pueden nombrar cobros sin nombre.'
        };
      }
    }

    // Escritura mínima: SOLO la celda del paciente de cada línea de la OP.
    for (var w = 0; w < filas.length; w++) sh.getRange(filas[w], IDENT_PAC_COL).setValue(nuevo);

    // Reemplazo quirúrgico del NOMBRE dentro de la observación bancaria.
    // No mueve montos, no borra filas, no recalcula saldos.
    var obsRes = _identSyncObsBanco(opId, nuevo);

    try {
      CacheService.getScriptCache().removeAll(['gas_ingresos_v1', 'gas_ingresos_v1_' + anio, 'gas_pacientes_v1']);
    } catch (eC) {}

    try {
      logAudit(usuario || 'sistema', 'Identificar pacientes', 'AsignarNombre', opId, 'Paciente',
        String(viejo).replace(/^\s+|\s+$/g, '') || '(vacío)', nuevo);
    } catch (eA) {}

    return { ok: true, op: opId, paciente: nuevo, lineas: filas.length, obsBanco: obsRes };
  } catch (ex) {
    return { ok: false, error: ex.message };
  } finally {
    try { lock.releaseLock(); } catch (eR) {}
  }
}

/* Actualiza el TEXTO del nombre en la observación de las filas bancarias de
   esta OP. La obs se escribe como  'Px. <paciente> [OP-00123]'  (ver
   updateIngresoConBancos en finance.gs), así que la fila se localiza por el
   tag [OP-…] y se sustituye solo el tramo del nombre.
   Si no hay fila, NO es error: hay cobros sin movimiento bancario. */
function _identSyncObsBanco(opId, nuevo) {
  var tocadas = 0;
  var claves = ['mercadopago', 'santander'];  // MP es el caso real; santander por completitud
  for (var c = 0; c < claves.length; c++) {
    try {
      if (typeof _bankSheetByKey !== 'function' || typeof _bankObsIdx !== 'function') break;
      var sheet = _bankSheetByKey(claves[c]);
      if (!sheet) continue;
      var lr = sheet.getLastRow();
      if (lr < 2) continue;
      var idx = _bankObsIdx(claves[c]);                 // mercadopago=8, santander=4
      var lc = Math.max(idx + 1, sheet.getLastColumn());
      var vals = sheet.getRange(2, 1, lr - 1, lc).getValues();
      var tag = '[' + opId + ']';
      for (var r = 0; r < vals.length; r++) {
        var obs = String(vals[r][idx] || '');
        if (obs.indexOf(tag) === -1) continue;
        // 'Px. <lo que sea> [OP-…]'  →  'Px. <nuevo> [OP-…]'
        var re = new RegExp('Px\\.\\s*[^\\[]*' + tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
        var next = obs.replace(re, 'Px. ' + nuevo + ' ' + tag);
        if (next !== obs) {
          sheet.getRange(r + 2, idx + 1).setValue(next);  // SOLO esa celda
          tocadas++;
        }
      }
    } catch (e) { /* la obs es cosmética: nunca debe tumbar la identificación */ }
  }
  return tocadas;
}

/* ══════════════════════════════════════════════════════════════════════
   3) SIMILARES — antipático de duplicados (endpoint público).
   ══════════════════════════════════════════════════════════════════════ */
function buscarPacientesSimilares(body) {
  try {
    // Acepta objeto {nombre, token} o string suelto (uso desde el editor).
    var esObj = body && typeof body === 'object';
    var nombre = esObj ? String(body.nombre || '') : String(body || '');
    if (esObj && !_tokenHasPermission(body.token || '', 'identificar_pacientes'))
      return { ok: false, error: 'Sin autorización (identificar_pacientes).', matches: [] };
    nombre = nombre.replace(/^\s+|\s+$/g, '');
    if (!nombre) return { ok: true, matches: [], umbral: IDENT_SIM_UMBRAL };
    var min = esObj && body.minScore !== undefined ? parseInt(body.minScore, 10) : 60;
    if (isNaN(min)) min = 60;
    return { ok: true, matches: _identSimilares(nombre, min), umbral: IDENT_SIM_UMBRAL };
  } catch (ex) {
    return { ok: false, error: ex.message, matches: [] };
  }
}

/* ══════════════════════════════════════════════════════════════════════
   4) ALTA RÁPIDA — dar de alta a la paciente sin salir del flujo.
   La similitud se RE-VALIDA AQUÍ (no se confía en el frontend): si hay un
   match >= IDENT_SIM_UMBRAL y no viene forzar:true, se rechaza devolviendo
   los candidatos para que la usuaria decida.
   Con forzar:true SÍ da de alta: dos pacientes pueden tener casi los mismos
   apellidos. Duplicar debe ser posible tras confirmar — nunca se bloquea en
   duro.
   ══════════════════════════════════════════════════════════════════════ */
function altaPacienteRapida(body) {
  body = body || {};
  if (!_tokenHasPermission(body.token || '', 'identificar_pacientes'))
    return { ok: false, error: 'Sin autorización (identificar_pacientes).' };

  var nombre = String(body.nombre || '').replace(/\s+/g, ' ').replace(/^\s+|\s+$/g, '');
  var email = String(body.email || '').replace(/^\s+|\s+$/g, '');
  var forzar = (body.forzar === true || String(body.forzar) === 'true');
  var usuario = body.usuario || '';
  if (!nombre) return { ok: false, error: 'Falta el nombre.' };
  if (_identEsNoLocalizado(nombre))
    return { ok: false, error: 'Ese no es un nombre real. Escribe el nombre de la paciente.' };

  // Re-validación en servidor: el frontend puede mentir u omitirse.
  var candidatos = _identSimilares(nombre, IDENT_SIM_UMBRAL);
  if (candidatos.length && !forzar) {
    return {
      ok: false, needConfirm: true, candidatos: candidatos, umbral: IDENT_SIM_UMBRAL,
      error: 'Ya existe una paciente muy parecida. Confirma si es la misma persona.'
    };
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (eLk) { return { ok: false, error: 'El sistema está ocupado. Intenta de nuevo en unos segundos.' }; }

  try {
    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    for (var h = 0; h < hdrs.length; h++) hdrs[h] = String(hdrs[h]).replace(/^\s+|\s+$/g, '');

    // Duplicado EXACTO: ni con forzar. Sería la misma fila dos veces, no un homónimo.
    var lr0 = sh.getLastRow();
    if (lr0 > 1) {
      var nomCol = sh.getRange(2, 2, lr0 - 1, 1).getValues();
      var nNorm = _identNorm(nombre);
      for (var d = 0; d < nomCol.length; d++) {
        if (_identNorm(nomCol[d][0]) === nNorm && nNorm)
          return { ok: false, error: 'Ya existe una paciente con ese nombre exacto: "' +
                   String(nomCol[d][0]).replace(/^\s+|\s+$/g, '') + '".' };
      }
    }

    var nuevoId = _identNextPacId(sh);
    var ancho = Math.max(hdrs.length, PAC_COL_LISTA);
    var fila = [];
    for (var z = 0; z < ancho; z++) fila.push('');
    fila[0] = nuevoId;                 // A = ID
    fila[1] = nombre;                  // B = Nombre
    if (email) fila[3] = email;        // D = Email
    fila[PAC_COL_LISTA - 1] = 'General';  // J = Lista de Precios (default del sistema)

    // Campos resueltos por header (no por posición): solo los que existan.
    var iAlta = _pacColIdx(hdrs, 'Fecha de Alta');
    if (iAlta === -1) iAlta = _pacColIdx(hdrs, 'Fecha Alta');
    if (iAlta > -1 && iAlta < ancho) fila[iAlta] = new Date();

    sh.appendRow(fila);

    try { CacheService.getScriptCache().removeAll(['gas_pacientes_v1']); } catch (eC) {}
    try {
      logAudit(usuario || 'sistema', 'Identificar pacientes', 'AltaRapida', nuevoId || nombre, 'Paciente',
        '—', nombre + (forzar ? ' (confirmado como persona distinta)' : ''));
    } catch (eA) {}

    return { ok: true, id: nuevoId, nombre: nombre, email: email, forzado: forzar };
  } catch (ex) {
    return { ok: false, error: ex.message };
  } finally {
    try { lock.releaseLock(); } catch (eR) {}
  }
}

/* Siguiente ID del catálogo de pacientes.
   Usa el patrón de _maxIdNum (finance.gs): BARRE toda la columna buscando el
   MÁXIMO real. Leer la última fila daba folios duplicados en producción
   (filas borradas/reordenadas, última fila con ID vacío) — no repetir ese bug.
   El prefijo y el ancho del cero-padding se deducen del catálogo mismo, para
   no inventar un formato que no sea el que ya usa la hoja. */
function _identNextPacId(sh) {
  var lr = sh.getLastRow();
  if (lr < 2) return '';
  var vals = sh.getRange(2, 1, lr - 1, 1).getValues();
  var prefCount = {}, vistos = 0, i, v, m;

  for (i = 0; i < vals.length; i++) {
    v = String(vals[i][0] || '').replace(/^\s+|\s+$/g, '');
    if (!v) continue;
    vistos++;
    m = v.match(/^([^0-9]*)(\d+)$/);
    if (!m) continue;
    prefCount[m[1]] = (prefCount[m[1]] || 0) + 1;
  }
  // Catálogo vacío o con IDs no numéricos: mejor dejar la celda en blanco que
  // inventar un formato. La hoja/operación lo resuelve como hoy.
  if (!vistos) return '';

  var bestPref = null, bestN = 0, p;
  for (p in prefCount) {
    if (prefCount.hasOwnProperty(p) && prefCount[p] > bestN) { bestN = prefCount[p]; bestPref = p; }
  }
  if (bestPref === null) return '';

  var esc = bestPref.replace(/[.*+?^${}()|[\]\\\-]/g, '\\$&');
  var re = new RegExp('^' + esc + '(\\d+)$');
  var max = _maxIdNum(sh, re);          // ← máximo REAL de toda la columna

  var width = 0;
  for (i = 0; i < vals.length; i++) {
    m = String(vals[i][0] || '').replace(/^\s+|\s+$/g, '').match(re);
    if (m && m[1].length > width) width = m[1].length;
  }
  var next = String(max + 1);
  while (next.length < width) next = '0' + next;
  return bestPref + next;
}
