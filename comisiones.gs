/* ═══════════════════════════════════════════════════════════════════════════
   COMISIONES POR VOLUMEN DE GRUPO (rebate a médicos externos + coordinador)
   ───────────────────────────────────────────────────────────────────────────
   LA REGLA DE NEGOCIO (textual del usuario, ya aclarada):
     "El Dr. Paladino llega con una paciente, Diana García. Ella paga su tratamiento
      directamente: Captura + Vitrificación, $23,600, precio GRUPO MEDICO. Si hacen
      entre el grupo de doctores 3 o más procedimientos en el mes les corresponde el
      10%. Ese 10% se le aplica solo al procedimiento que vaya a hacer el Dr., pero
      para llegar sí se suman todos los involucrados. Daniel Madero no hace
      procedimientos: a él se le paga una comisión del 5% de ese 10%."

   EL MODELO (4 pasos):
     1. CONTAR   — todos los procedimientos ELEGIBLES (lista configurable de
                   productos) de TODO el grupo en el mes. El grupo desbloquea.
     2. ESCALON  — la cantidad determina el nivel (progresivo: 1-2 = 0%, 3+ = 10%…).
     3. BASE     — el % del escalón se aplica SOLO a los procedimientos de ESE
                   médico. Paladino: $23,600 × 10% = $2,360. Cada quien cobra lo suyo.
     4. REPARTO  — mitad y mitad: $1,180 de crédito al Dr. Paladino (contra SU
                   cuenta) y $1,180 en efectivo a Daniel Madero. Ambos el mes
                   siguiente (`diferido`).

   POR QUE `parte` ES EL % DEL ESCALON Y NO DEL MONTO:
     El usuario dijo "5% de ese 10%". Leído literal eso sería 0.5% — no es lo que
     quiere. `parte:50` sobre un escalón de 10% da el 5% efectivo que él describe.
     Guardarlo así hace que un futuro escalón de 15% siga repartiéndose mitad y
     mitad SIN reconfigurar nada. El panel muestra el % efectivo resultante para
     que el número real quede a la vista antes de aprobar.

   FUENTE DE DATOS — UNA SOLA: BD_Ingresos + la columna OrigenExterno.
     Los dos casos del usuario aterrizan ahí:
       · La paciente paga directo  → venta a nombre de la paciente, OrigenExterno = el médico.
       · Al médico se le factura   → cargarSaldoInicial llama saveIngreso, que también
                                     crea la venta en BD_Ingresos con el médico de titular.
     (El módulo de "Descuentos por agencia" NO sirve para esto y no se toca: lee solo
      de Cuentas_Cobrar, no ve una venta pagada directo, no escribe nada y no tiene
      beneficiario — el descuento siempre baja al titular del cargo.)

   ESTE ARCHIVO NO ESCRIBE NADA POR SU CUENTA: calcularComisiones() es SOLO LECTURA.
   Solo generarComisiones() escribe, y exige permiso + aprobación explícita del
   usuario en pantalla. Es dinero, y un error se multiplica por todos los médicos.

   Requiere: origenes.gs (_origResolver), cobranza.gs (_cobTierPct, crédito a favor),
             finance.gs (saveCxP, libros por año, permisos, logAudit).
   ═══════════════════════════════════════════════════════════════════════════ */

var COMISIONES_VER     = 'comisiones-2026.07.16a';
var COMISIONES_CFG_KEY = 'COMISIONES_CFG';
var COMISIONES_TAB     = 'Comisiones_Generadas';
var COMISIONES_PERM    = 'generar_comisiones';

/* ───────────────────────── Helpers propios ─────────────────────────
   Se apoyan en cobranza.gs/finance.gs donde ya existe la pieza (_cobNum, _cobTierPct,
   _cobKeyNom, _cobKeyProd, _ingEsCancelada…). Aquí solo va lo que NO existe. */

function _comNum(v) { return (typeof _cobNum === 'function') ? _cobNum(v) : (parseFloat(String(v || '').replace(/[$,\s]/g, '')) || 0); }
function _comDeacento(s) {
  return String(s || '').replace(/[áàäâ]/g, 'a').replace(/[éèëê]/g, 'e').replace(/[íìïî]/g, 'i')
    .replace(/[óòöô]/g, 'o').replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n');
}
/* Clave de producto: _cobKeyProd (minúsculas + espacios normalizados) PERO además sin
   acentos. _cobKeyProd solo no basta aquí: si el catálogo dice "Captura + Vitrificación"
   y alguien escribe "Vitrificacion" en la config, un match exacto fallaría en silencio
   y el procedimiento NO contaría — dinero mal calculado sin un solo error visible. */
function _comKeyProd(s) {
  var base = (typeof _cobKeyProd === 'function') ? _cobKeyProd(s) : String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  return _comDeacento(base);
}
function _comKeyNom(s) {
  return (typeof _cobKeyNom === 'function') ? _cobKeyNom(s)
       : _comDeacento(String(s || '').trim().toLowerCase()).replace(/\s+/g, ' ').trim();
}
function _comPad2(n) { return ('0' + n).slice(-2); }
/* Suma n meses a 'yyyy-MM' (n negativo = hacia atrás). cobranza.gs tiene _cobMesPrev,
   pero está clavado en -1 y aquí se necesita ir HACIA ADELANTE (el pago es el mes
   siguiente, `diferido`). Esta función es su superconjunto exacto:
   _comMesSuma(m,-1) === _cobMesPrev(m) — y la prueba lo verifica, para que no diverjan. */
function _comMesSuma(mes, n) {
  var y = parseInt(String(mes).substring(0, 4), 10), m = parseInt(String(mes).substring(5, 7), 10);
  if (!y || !m) return mes;
  var t = (y * 12) + (m - 1) + (parseInt(n, 10) || 0);
  return Math.floor(t / 12) + '-' + _comPad2((t % 12) + 1);
}
function _comMesKey(anio, mes) { return parseInt(anio, 10) + '-' + _comPad2(parseInt(mes, 10)); }
/* Último día del mes 'yyyy-MM' → 'yyyy-MM-dd'. El vencimiento de la CxP: la comisión
   se paga DURANTE el mes siguiente, así que vence al cierre de ese mes. */
function _comFinDeMes(mes) {
  var y = parseInt(String(mes).substring(0, 4), 10), m = parseInt(String(mes).substring(5, 7), 10);
  return String(mes) + '-' + _comPad2(new Date(y, m, 0).getDate());
}
function _comRedondea(n) { return Math.round((_comNum(n) + Number.EPSILON) * 100) / 100; }

/* ── Guard de deploy parcial. Apps Script comparte scope global: si origenes.gs no
   está desplegado (o está en una versión vieja), _origResolver no existe y las ventas
   NO se podrían agrupar. Reventar aquí es lo correcto: seguir sin él calcularía cero
   comisiones y parecería "este mes no hubo" en vez de "el backend está a medias". ── */
function _comGuardOrig() {
  if (typeof _origResolver !== 'function')
    return 'Actualiza origenes.gs en Apps Script y redespliega (falta _origResolver): sin la jerarquía de orígenes no se puede saber qué médico pertenece a qué grupo, y las comisiones saldrían todas en cero.';
  return '';
}
function _comGuardDeps() {
  var g = _comGuardOrig(); if (g) return g;
  if (typeof _cobTierPct !== 'function')
    return 'Actualiza cobranza.gs en Apps Script y redespliega (falta _cobTierPct): sin los escalones no se puede calcular el nivel del grupo.';
  if (typeof _cobRegistrarCreditoFavor !== 'function')
    return 'Actualiza cobranza.gs en Apps Script y redespliega (falta _cobRegistrarCreditoFavor): sin él no se puede acreditar al médico.';
  if (typeof saveCxP !== 'function' || typeof INGRESOS_IDS === 'undefined')
    return 'Actualiza finance.gs en Apps Script y redespliega (faltan saveCxP / los libros de ingresos).';
  return '';
}

/* ═════════════════════════ CONFIG (Script Property) ═════════════════════════
 * COMISIONES_CFG = { reglas:[{
 *     id, nombre, activo, grupoId,          // el grupo cuyo VOLUMEN cuenta
 *     productos:[...],                      // SOLO estos cuentan Y generan
 *     escalones:[{desde,hasta,pct}],        // progresivos; hasta:'' = sin tope
 *     beneficiarios:[{ tipo:'medico_del_procedimiento'|'fijo', origenId, parte, via }],
 *     diferido                              // meses: 1 = se aplica/paga el mes N+1
 * }] }
 * `parte` = % DEL ESCALON (suma de partes debe dar 100). `via` = 'nota_credito'|'efectivo'.
 * Sin config previa arranca vacío: el usuario captura su primera regla desde el panel.
 */
var COMISIONES_VIAS  = ['nota_credito', 'efectivo'];   // enum abierto: agregar aquí
var COMISIONES_TIPOS = ['medico_del_procedimiento', 'fijo'];

function _comCfg() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(COMISIONES_CFG_KEY);
    if (raw) { var o = JSON.parse(raw); if (o && o.reglas) return o; }
  } catch (e) {}
  return { reglas: [] };
}
function _comRegla(reglaId) {
  var found = null;
  (_comCfg().reglas || []).forEach(function (r) { if (String(r.id) === String(reglaId)) found = r; });
  return found;
}
/* Escalones normalizados para _cobTierPct: la función de cobranza compara
   `qty <= tiers[i].hasta` sin normalizar, así que un `hasta:''` la haría devolver 0
   siempre (el tope abierto quedaría cerrado). Se traduce a 1e9, igual que hace
   _cobDescuentosAgencia antes de llamarla. */
function _comTiers(regla) {
  return (regla.escalones || []).map(function (t) {
    return { desde: _comNum(t.desde) || 1,
             hasta: (t.hasta === '' || t.hasta == null) ? 1e9 : (_comNum(t.hasta) || 1e9),
             pct: _comNum(t.pct) };
  }).sort(function (a, b) { return a.desde - b.desde; });
}

/* Valida una regla. Devuelve [] si está bien, o la lista de problemas.
   La suma de `parte` DEBE dar 100: si diera menos, la clínica se quedaría con un
   pedazo del rebate que ya prometió; si diera más, pagaría de más. Ninguna de las
   dos puede pasar en silencio. */
function _comValidarRegla(r) {
  var errs = [];
  if (!r || !String(r.id || '').trim()) errs.push('La regla necesita un id.');
  if (!String(r.grupoId || '').trim()) errs.push('Falta el grupo (grupoId) cuyo volumen cuenta.');
  if (!(r.productos || []).length) errs.push('Falta la lista de procedimientos que cuentan (sin ella no contaría nada).');
  if (!(r.escalones || []).length) errs.push('Faltan los escalones (desde / hasta / %).');
  var bens = r.beneficiarios || [];
  if (!bens.length) errs.push('Faltan los beneficiarios del reparto.');
  var suma = 0;
  bens.forEach(function (b, i) {
    suma += _comNum(b.parte);
    if (COMISIONES_TIPOS.indexOf(String(b.tipo)) < 0) errs.push('Beneficiario ' + (i + 1) + ': tipo desconocido "' + b.tipo + '".');
    if (COMISIONES_VIAS.indexOf(String(b.via)) < 0)   errs.push('Beneficiario ' + (i + 1) + ': vía de pago desconocida "' + b.via + '".');
    if (String(b.tipo) === 'fijo' && !String(b.origenId || '').trim())
      errs.push('Beneficiario ' + (i + 1) + ': un beneficiario fijo necesita su origenId del catálogo.');
  });
  if (bens.length && Math.abs(suma - 100) > 0.01)
    errs.push('El reparto suma ' + suma + '% y debe sumar exactamente 100% (es el reparto DEL escalón, no del monto).');
  var tiers = _comTiers(r || {});
  for (var i = 1; i < tiers.length; i++)
    if (tiers[i].desde <= tiers[i - 1].hasta && tiers[i - 1].hasta !== 1e9)
      errs.push('Los escalones se enciman entre ' + tiers[i - 1].desde + '-' + tiers[i - 1].hasta + ' y ' + tiers[i].desde + '-' + (tiers[i].hasta === 1e9 ? '∞' : tiers[i].hasta) + '.');
  return errs;
}

/* Preview del reparto para el panel: traduce `parte` a % EFECTIVO sobre la venta.
   "Del 10% del escalón: 50% al médico = 5% · 50% a Madero = 5%" — el usuario dijo
   "5% de ese 10%", que leído literal sería 0.5%. Este preview desambigua ANTES de
   guardar la regla. */
function _comPreviewReparto(regla) {
  var tiers = _comTiers(regla || {});
  var out = [];
  tiers.forEach(function (t) {
    if (t.pct <= 0) return;
    (regla.beneficiarios || []).forEach(function (b) {
      out.push({ escalonPct: t.pct, desde: t.desde, hasta: t.hasta === 1e9 ? '' : t.hasta,
                 tipo: b.tipo, origenId: b.origenId || '', via: b.via,
                 parte: _comNum(b.parte),
                 efectivoPct: _comRedondea(t.pct * _comNum(b.parte) / 100) });
    });
  });
  return out;
}

function readComisionesCfg() {
  try {
    var cfg = _comCfg();
    var reglas = (cfg.reglas || []).map(function (r) {
      var c = JSON.parse(JSON.stringify(r));
      c._preview = _comPreviewReparto(r);
      c._errores = _comValidarRegla(r);
      return c;
    });
    return { ok: true, version: COMISIONES_VER, reglas: reglas, vias: COMISIONES_VIAS, tipos: COMISIONES_TIPOS,
             deploy: _comGuardDeps() || '' };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

function saveComisionesCfg(body) {
  try {
    body = body || {};
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token, COMISIONES_PERM))
      return { ok: false, error: 'No tienes permiso para configurar comisiones (' + COMISIONES_PERM + ').', version: COMISIONES_VER };
    var list = body.reglas || [];
    if (!Array.isArray(list)) return { ok: false, error: 'Formato inválido: se esperaba una lista de reglas.', version: COMISIONES_VER };
    // Validar TODAS antes de guardar ninguna: media config es peor que ninguna.
    var problemas = [];
    list.forEach(function (r) {
      _comValidarRegla(r).forEach(function (e) { problemas.push((r && r.nombre ? r.nombre : (r && r.id) || '(sin nombre)') + ': ' + e); });
    });
    if (problemas.length) return { ok: false, error: 'La configuración tiene errores:\n· ' + problemas.join('\n· '), problemas: problemas, version: COMISIONES_VER };
    var limpio = list.map(function (r) {
      return { id: String(r.id).trim(), nombre: String(r.nombre || '').trim(), activo: r.activo !== false,
               grupoId: String(r.grupoId).trim(),
               productos: (r.productos || []).map(function (p) { return String(p).trim(); }).filter(function (p) { return !!p; }),
               escalones: (r.escalones || []).map(function (t) { return { desde: _comNum(t.desde) || 1, hasta: (t.hasta === '' || t.hasta == null) ? '' : _comNum(t.hasta), pct: _comNum(t.pct) }; }),
               beneficiarios: (r.beneficiarios || []).map(function (b) { return { tipo: String(b.tipo), origenId: String(b.origenId || '').trim(), parte: _comNum(b.parte), via: String(b.via) }; }),
               diferido: (r.diferido == null || r.diferido === '') ? 1 : (parseInt(r.diferido, 10) || 0),
               notas: String(r.notas || '') };
    });
    PropertiesService.getScriptProperties().setProperty(COMISIONES_CFG_KEY, JSON.stringify({ reglas: limpio }));
    try { logAudit(body.usuario || '', 'Comisiones', 'Guardar configuración', '', '', '', limpio.length + ' regla(s)'); } catch (e) {}
    return { ok: true, version: COMISIONES_VER, guardadas: limpio.length };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═════════════════════════ HOJA DE CONTROL ═════════════════════════
 * Comisiones_Generadas es la MEMORIA de lo ya pagado — la única defensa contra
 * generar el mismo mes dos veces. Vive en INGRESOS_SS_ID junto a Creditos_Favor.
 * Esquema: Mes | ReglaId | Beneficiario | Via | Monto | RefId | Estado | Usuario | Timestamp
 *   RefId  = OP sintética (crédito) o el ID de la fila de Egresos (CxP) → permite
 *            volver a encontrar lo escrito para revisarlo o revertirlo.
 *   Estado = 'activa' | 'revertida'
 */
function _comEnsureControl() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COMISIONES_TAB);
  if (!sh) {
    sh = ss.insertSheet(COMISIONES_TAB);
    sh.appendRow(['Mes', 'ReglaId', 'Beneficiario', 'Via', 'Monto', 'RefId', 'Estado', 'Usuario', 'Timestamp']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f3f4f6');
  }
  return sh;
}
/* Lo ya generado (y NO revertido) para un mes+regla. Vacío = nunca se ha generado. */
function _comYaGenerado(mes, reglaId) {
  var out = [];
  try {
    var sh = _comEnsureControl();
    var raw = sh.getDataRange().getValues();
    for (var i = 1; i < raw.length; i++) {
      if (String(raw[i][0] || '').trim() !== String(mes)) continue;
      if (String(raw[i][1] || '').trim() !== String(reglaId)) continue;
      if (String(raw[i][6] || '').trim().toLowerCase() === 'revertida') continue;
      out.push({ rowNum: i + 1, mes: String(raw[i][0]), reglaId: String(raw[i][1]),
                 beneficiario: String(raw[i][2]), via: String(raw[i][3]),
                 monto: _comNum(raw[i][4]), refId: String(raw[i][5] || '').trim(),
                 estado: String(raw[i][6] || ''), usuario: String(raw[i][7] || ''),
                 timestamp: raw[i][8] });
    }
  } catch (e) {}
  return out;
}

/* OP sintética del crédito: estable y única por (regla, mes, médico). _cobRegistrarCreditoFavor
   hace UPSERT POR OP, así que esta llave es lo que impide que regenerar apile créditos. */
function _comOpCredito(reglaId, mes, origenId) { return 'COM-' + reglaId + '-' + mes + '-' + origenId; }

/* ═════════════════════════ MOTOR (SOLO LECTURA) ═════════════════════════ */

/* Lee BD_Ingresos del año pedido. Columnas SIEMPRE por encabezado: OrigenExterno es
   una columna dinámica (_ingColEnsure la crea al final cuando falta), así que leerla
   por índice fijo es un bug esperando su turno. */
function _comReadVentas(anio, mesKey) {
  // Año sin libro → RECHAZAR. _ingIdDeAnio cae al libro 2026 cuando no conoce el año:
  // pedir 2027 devolvería datos de 2026 y calcularía comisiones sobre el mes equivocado
  // sin decir una palabra. Preferimos el error.
  if (typeof INGRESOS_IDS === 'undefined' || !INGRESOS_IDS[parseInt(anio, 10)])
    return { ok: false, error: 'No hay libro de ingresos para ' + anio + '. Los libros configurados son: ' + Object.keys(INGRESOS_IDS || {}).join(', ') + '.' };
  var ssId = _ingIdDeAnio(anio);
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName(typeof BD_INGRESOS_TAB !== 'undefined' ? BD_INGRESOS_TAB : 'BD_Ingresos');
  if (!sh) return { ok: false, error: 'No se encontró BD_Ingresos en el libro de ' + anio + '.' };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, rows: [], nombresPaciente: {} };

  var hdr = data[0].map(function (c) { return _comDeacento(String(c).trim().toLowerCase()); });
  function col(keys, fb) {
    for (var n = 0; n < keys.length; n++) for (var c = 0; c < hdr.length; c++) if (hdr[c] === keys[n]) return c;
    for (var n2 = 0; n2 < keys.length; n2++) for (var c2 = 0; c2 < hdr.length; c2++) if (hdr[c2].indexOf(keys[n2]) > -1) return c2;
    return fb;
  }
  var iOp = col(['op'], 0), iFecha = col(['fecha'], 2), iPac = col(['paciente'], 3),
      iProd = col(['producto'], 5), iCant = col(['cantidad'], 8),
      iTot = col(['totalpagar', 'total a pagar', 'total'], 9),
      iOrigen = col(['origenexterno', 'origen'], -1),
      // Sin fallback: si la hoja del año no trae la columna, nada está cancelado.
      iCancel = col(['cancelada'], -1);

  var rows = [], nombresPaciente = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var op = String(r[iOp] || '').trim(); if (!op) continue;
    // Todo nombre de titular visto, para avisar si el crédito de un médico quedaría huérfano.
    if (String(r[iPac] || '').trim()) nombresPaciente[_comKeyNom(r[iPac])] = String(r[iPac]).trim();
    // VENTA CANCELADA → fuera. Una comisión sobre una venta cancelada es dinero regalado:
    // la venta se reversó en bancos pero la fila sigue en la hoja con su TotalPagar intacto.
    if (iCancel > -1 && typeof _ingEsCancelada === 'function' && _ingEsCancelada(r[iCancel])) continue;
    var d = (r[iFecha] instanceof Date) ? r[iFecha] : (typeof _cobD === 'function' ? _cobD(r[iFecha]) : new Date(r[iFecha]));
    if (!d || isNaN(d.getTime())) continue;
    var mk = d.getFullYear() + '-' + _comPad2(d.getMonth() + 1);
    if (mesKey && mk !== mesKey) continue;
    rows.push({ op: op, fecha: mk, paciente: String(r[iPac] || '').trim(),
                producto: String(r[iProd] || '').trim(),
                cantidad: _comNum(r[iCant]) || 1,
                total: _comNum(r[iTot]),
                origen: iOrigen > -1 ? String(r[iOrigen] || '').trim() : '' });
  }
  return { ok: true, rows: rows, nombresPaciente: nombresPaciente };
}

/* calcularComisiones({anio, mes, reglaId}) → SOLO LECTURA. No escribe absolutamente nada.
   Es lo que alimenta la pantalla de revisión previa: el usuario ve el detalle línea por
   línea y los totales por beneficiario ANTES de que exista un solo peso. */
function calcularComisiones(body) {
  try {
    body = body || {};
    var dep = _comGuardDeps(); if (dep) return { ok: false, error: dep, version: COMISIONES_VER };

    var anio = parseInt(body.anio, 10), mes = parseInt(body.mes, 10);
    if (!anio || !mes || mes < 1 || mes > 12) return { ok: false, error: 'Indica un año y un mes válidos.', version: COMISIONES_VER };
    var mesKey = _comMesKey(anio, mes);

    var regla = _comRegla(body.reglaId);
    if (!regla) return { ok: false, error: 'No existe la regla "' + body.reglaId + '". Configúrala en Panel de Control → Comisiones.', version: COMISIONES_VER };
    if (regla.activo === false) return { ok: false, error: 'La regla "' + (regla.nombre || regla.id) + '" está desactivada.', version: COMISIONES_VER };
    var errs = _comValidarRegla(regla);
    if (errs.length) return { ok: false, error: 'La regla tiene errores de configuración:\n· ' + errs.join('\n· '), problemas: errs, version: COMISIONES_VER };

    var lect = _comReadVentas(anio, mesKey);
    if (!lect.ok) return { ok: false, error: lect.error, version: COMISIONES_VER };

    var prodSet = {};
    (regla.productos || []).forEach(function (p) { prodSet[_comKeyProd(p)] = 1; });

    // ── 1) Ventas del GRUPO con producto elegible ──────────────────────────
    var elegibles = [], descartadas = { sinOrigen: 0, otroGrupo: 0, productoNoElegible: 0 };
    // Detalle SOLO de las que SÍ son del grupo pero su PRODUCTO no está en la lista
    // elegible: es lo accionable ("Sasha con Paladino no cuenta porque Inseminación
    // no está en los productos de la regla"). sinOrigen/otroGrupo solo se cuentan.
    var descNoElegible = [];
    lect.rows.forEach(function (v) {
      if (!v.origen) { descartadas.sinOrigen++; return; }
      var res = null;
      try { res = _origResolver(v.origen); } catch (e) { res = null; }
      if (!res || !res.id) { descartadas.sinOrigen++; return; }
      if (String(res.grupoId) !== String(regla.grupoId)) { descartadas.otroGrupo++; return; }
      if (!prodSet[_comKeyProd(v.producto)]) {
        descartadas.productoNoElegible++;
        descNoElegible.push({ op: v.op, paciente: v.paciente, producto: v.producto,
                              cantidad: v.cantidad, total: v.total, medicoNombre: res.nombre });
        return;
      }
      elegibles.push({ op: v.op, paciente: v.paciente, producto: v.producto, cantidad: v.cantidad,
                       total: v.total, medicoId: res.id, medicoNombre: res.nombre, medicoTipo: res.tipo });
    });

    // ── 2) Conteo del grupo → escalón ───────────────────────────────────────
    // Se cuentan PROCEDIMIENTOS, no renglones: una venta con cantidad 2 son 2 procedimientos.
    var conteo = 0;
    elegibles.forEach(function (e) { conteo += e.cantidad; });
    var tiers = _comTiers(regla);
    var pct = _cobTierPct(tiers, conteo);
    var escalon = null;
    tiers.forEach(function (t) { if (conteo >= t.desde && conteo <= t.hasta) escalon = t; });

    // ── 3) Base POR MÉDICO (el grupo desbloquea; cada quien cobra lo suyo) ──
    var porMedico = {};
    elegibles.forEach(function (e) {
      if (!porMedico[e.medicoId]) porMedico[e.medicoId] = { medicoId: e.medicoId, medicoNombre: e.medicoNombre, base: 0, cantidad: 0, ventas: [] };
      porMedico[e.medicoId].base += e.total;
      porMedico[e.medicoId].cantidad += e.cantidad;
      porMedico[e.medicoId].ventas.push(e);
    });

    // ── 4) Reparto entre beneficiarios ──────────────────────────────────────
    var mesAplica = _comMesSuma(mesKey, (regla.diferido == null ? 1 : parseInt(regla.diferido, 10) || 0));
    var detalle = [], totales = {}, avisos = [];

    if (pct > 0) {
      Object.keys(porMedico).forEach(function (mid) {
        var M = porMedico[mid];
        var accrual = _comRedondea(M.base * pct / 100);
        (regla.beneficiarios || []).forEach(function (b) {
          var monto = _comRedondea(accrual * _comNum(b.parte) / 100);
          if (monto <= 0.01) return;
          var benId = '', benNombre = '';
          if (String(b.tipo) === 'medico_del_procedimiento') { benId = M.medicoId; benNombre = M.medicoNombre; }
          else {
            var rf = null; try { rf = _origResolver(b.origenId); } catch (e) { rf = null; }
            if (!rf || !rf.id) { avisos.push('El beneficiario fijo "' + b.origenId + '" no existe en el catálogo de orígenes: su parte no se generó.'); return; }
            benId = rf.id; benNombre = rf.nombre;
          }
          detalle.push({ medicoId: M.medicoId, medicoNombre: M.medicoNombre,
                         base: _comRedondea(M.base), cantidadMedico: M.cantidad,
                         pct: pct, accrual: accrual,
                         beneficiarioId: benId, beneficiario: benNombre,
                         tipo: b.tipo, parte: _comNum(b.parte),
                         efectivoPct: _comRedondea(pct * _comNum(b.parte) / 100),
                         via: b.via, monto: monto, mesAplica: mesAplica,
                         ventas: M.ventas.map(function (v) { return { op: v.op, paciente: v.paciente, producto: v.producto, cantidad: v.cantidad, total: v.total }; }) });
          // TOTALES POR BENEFICIARIO = lo que de verdad se escribe. Un beneficiario fijo
          // (el coordinador) cobra por los procedimientos de VARIOS médicos: si se
          // escribiera una CxP por cada médico, Madero tendría 3 pagos sueltos el mismo
          // mes en vez del pago único que describe la regla. El detalle de arriba es para
          // que se vea de dónde salió cada peso; esto es lo que se paga.
          var tk = benId + '|' + b.via;
          if (!totales[tk]) totales[tk] = { beneficiarioId: benId, beneficiario: benNombre, via: b.via, monto: 0, mesAplica: mesAplica, lineas: [] };
          totales[tk].monto = _comRedondea(totales[tk].monto + monto);
          totales[tk].lineas.push({ medicoNombre: M.medicoNombre, base: _comRedondea(M.base), pct: pct, parte: _comNum(b.parte), efectivoPct: _comRedondea(pct * _comNum(b.parte) / 100), monto: monto });
        });
      });
    }

    // ── Avisos honestos sobre las vías de pago ──────────────────────────────
    // (a) El crédito solo se puede consumir contra una cuenta cuyo titular normalizado
    //     sea IDÉNTICO. Si el médico nunca ha sido titular de una venta, el crédito
    //     nace y se queda ahí. Mejor avisarlo antes de crearlo que después.
    detalle.forEach(function (d) {
      if (d.via !== 'nota_credito') return;
      if (!lect.nombresPaciente[_comKeyNom(d.beneficiario)])
        avisos.push('El crédito de ' + d.beneficiario + ' ($' + d.monto.toLocaleString('es-MX') + ') quedará a su nombre, pero en ' + anio + ' no aparece ninguna cuenta a nombre de "' + d.beneficiario + '". Solo se podrá aplicar cuando se le facture con ESE mismo nombre.');
    });
    // (b) saveCxP escribe SIEMPRE en el libro Egresos2026 (está clavado ahí, no es
    //     year-aware). Si la comisión se paga en otro año, la CxP caería en el libro
    //     equivocado. No se genera a ciegas: se avisa aquí y se BLOQUEA al generar.
    var anioPago = parseInt(String(mesAplica).substring(0, 4), 10);
    var bloqueoLibro = '';
    var hayEfectivo = detalle.some(function (d) { return d.via === 'efectivo'; });
    if (hayEfectivo && typeof EGRESOS_IDS !== 'undefined' && typeof EGRESOS_SS_2026 !== 'undefined' && EGRESOS_IDS[anioPago] !== EGRESOS_SS_2026) {
      bloqueoLibro = 'La comisión se paga en ' + mesAplica + ', pero las cuentas por pagar se escriben siempre en el libro Egresos2026. Generarla dejaría el pago a ' + (detalle.filter(function (d) { return d.via === 'efectivo'; })[0] || {}).beneficiario + ' en el libro del año equivocado. Antes de generar: agrega el libro de ' + anioPago + ' y haz saveCxP year-aware.';
      avisos.push(bloqueoLibro);
    }

    var ya = _comYaGenerado(mesKey, regla.id);
    var totLista = Object.keys(totales).map(function (k) { return totales[k]; })
      .sort(function (a, b) { return b.monto - a.monto; });

    return { ok: true, version: COMISIONES_VER,
             mes: mesKey, mesAplica: mesAplica, anio: anio,
             regla: { id: regla.id, nombre: regla.nombre, diferido: regla.diferido == null ? 1 : regla.diferido, productos: regla.productos },
             conteo: conteo, escalon: escalon ? { desde: escalon.desde, hasta: escalon.hasta === 1e9 ? '' : escalon.hasta, pct: escalon.pct } : null,
             pct: pct, escalones: tiers.map(function (t) { return { desde: t.desde, hasta: t.hasta === 1e9 ? '' : t.hasta, pct: t.pct }; }),
             detalle: detalle, totales: totLista,
             totalGeneral: _comRedondea(totLista.reduce(function (s, t) { return s + t.monto; }, 0)),
             elegibles: elegibles, descartadas: descartadas, descartadasDetalle: descNoElegible,
             avisos: avisos, bloqueo: bloqueoLibro,
             yaGenerado: ya, generado: ya.length > 0 };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═══════════════ AVISO DE DESCUENTO/REBATE POR ORIGEN (SOLO LECTURA) ═══════════════
 * Para el aviso al CAPTURAR un ingreso: dado el origen elegido (médico externo o
 * agencia) y el mes, devuelve el % que aplica ESTE MES. NO escribe nada.
 *  · Médico externo → busca su grupo (_origResolver) → la regla de comisión activa
 *    de ese grupo → calcularComisiones → pct del escalón alcanzado por el volumen
 *    del grupo. (Es un REBATE al médico, no una rebaja al precio del paciente.)
 *  · Agencia (Reprovida) → _cobDescuentosAgencia → % de volumen del mes.
 * Devuelve {ok, tipo:'medico'|'agencia'|'', nombre, pct, conteo, escalones, ...}.
 */
function descuentoOrigen(origen, anio, mes) {
  try {
    origen = String(origen || '').trim();
    if (!origen) return { ok: true, tipo: '', pct: 0 };
    var now = new Date();
    anio = parseInt(anio, 10) || now.getFullYear();
    mes  = parseInt(mes, 10)  || (now.getMonth() + 1);

    // (A) MÉDICO EXTERNO → regla de su grupo → pct del mes.
    var res = null;
    try { if (typeof _origResolver === 'function') res = _origResolver(origen); } catch (e) {}
    if (res && res.id) {
      var grupoId = String(res.grupoId || res.id);
      var reglas = (_comCfg().reglas || []).filter(function (r) {
        return r.activo !== false && String(r.grupoId) === grupoId;
      });
      if (reglas.length) {
        // El descuento que APLICA a una venta de (mes) se GANÓ 'diferido' meses antes:
        // con diferido=1, una venta de julio usa el volumen acumulado de JUNIO (el mes
        // que ya cerró). Con diferido=0 usa el mes en curso. Por eso se calcula sobre el
        // mes GANADO = mes − diferido, no sobre el mes de la venta.
        var difer = (reglas[0].diferido == null) ? 1 : (parseInt(reglas[0].diferido, 10) || 0);
        var earnKey = _comMesSuma(_comMesKey(anio, mes), -difer);
        var eAnio = parseInt(String(earnKey).substring(0, 4), 10), eMes = parseInt(String(earnKey).substring(5, 7), 10);
        var calc = calcularComisiones({ anio: eAnio, mes: eMes, reglaId: reglas[0].id });
        if (calc && calc.ok) {
          return { ok: true, tipo: 'medico', nombre: res.nombre || origen,
                   regla: reglas[0].nombre || reglas[0].id,
                   pct: calc.pct || 0, conteo: calc.conteo || 0,
                   escalon: calc.escalon || null, escalones: calc.escalones || [],
                   diferido: difer, mesGanado: calc.mes, mesAplica: _comMesKey(anio, mes) };
        }
      }
    }
    // (B) AGENCIA → descuento por volumen del mes.
    try {
      if (typeof _cobDescuentosAgencia === 'function') {
        var desc = _cobDescuentosAgencia(origen);
        if (desc && desc.aplica) {
          var pctMax = 0;
          (desc.lineas || []).forEach(function (l) { if ((_comNum(l.pct) || 0) > pctMax) pctMax = _comNum(l.pct); });
          return { ok: true, tipo: 'agencia', nombre: origen, pct: pctMax, total: desc.total || 0 };
        }
      }
    } catch (e2) {}
    return { ok: true, tipo: '', pct: 0, nombre: origen };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ═══════════════ COMISIONES PENDIENTES DE GENERAR (SOLO LECTURA) ═══════════════
 * Recorre las reglas activas × los últimos meses cerrados y lista las que tienen
 * pct>0 y AÚN no se han generado. Alimenta el aviso "hay comisiones por pagar".
 * Solo lectura (calcularComisiones no escribe). `meses` = cuántos meses hacia atrás
 * revisar (default 3, incluyendo el mes anterior; el mes en curso no se cierra aún).
 */
function comisionesPendientes(meses) {
  try {
    meses = parseInt(meses, 10) || 3;
    var reglas = (_comCfg().reglas || []).filter(function (r) { return r.activo !== false; });
    if (!reglas.length) return { ok: true, pendientes: [], total: 0 };
    var now = new Date();
    var pend = [], totalMonto = 0;
    // Del mes anterior hacia atrás (el mes actual sigue acumulando, no se genera).
    for (var k = 1; k <= meses; k++) {
      var d = new Date(now.getFullYear(), now.getMonth() - k, 1);
      var anio = d.getFullYear(), mes = d.getMonth() + 1;
      for (var i = 0; i < reglas.length; i++) {
        var calc = null;
        try { calc = calcularComisiones({ anio: anio, mes: mes, reglaId: reglas[i].id }); } catch (e) { calc = null; }
        if (!calc || !calc.ok) continue;
        if ((calc.pct || 0) > 0 && !calc.generado && calc.totalGeneral > 0.01) {
          pend.push({ reglaId: reglas[i].id, regla: reglas[i].nombre || reglas[i].id,
                      anio: anio, mes: mes, mesKey: calc.mes, pct: calc.pct,
                      conteo: calc.conteo, total: calc.totalGeneral });
          totalMonto += calc.totalGeneral;
        }
      }
    }
    return { ok: true, pendientes: pend, total: _comRedondea(totalMonto) };
  } catch (ex) { return { ok: false, error: ex.message, pendientes: [] }; }
}

/* ═════════════════════════ REVERSA (para regenerar) ═════════════════════════
 * Solo se puede revertir lo que NADIE ha tocado todavía. Se revisa TODO antes de
 * revertir NADA: una reversa a medias dejaría al médico sin crédito y a Madero con
 * su CxP viva, o al revés.
 *
 * Por qué es indispensable revisar:
 *   · Creditos_Favor.MontoCredito es el SALDO VIVO — consumirlo lo decrementa. Y
 *     _cobRegistrarCreditoFavor hace upsert ABSOLUTO. Regenerar a ciegas sobre un
 *     crédito ya gastado a medias lo REGRESARÍA a su monto original: le devolvería
 *     al médico un dinero que ya se gastó. Dinero duplicado, en silencio.
 *   · La CxP en Egresos puede estar YA PAGADA: revertirla borraría el registro de
 *     un pago que de verdad salió de la caja.
 */
function _comEstadoRef(item) {
  // → { tocado:bool, motivo:'', rowNum } para un renglón ya generado.
  try {
    if (item.via === 'nota_credito') {
      var sh = _cobEnsureCreditos();
      var raw = sh.getDataRange().getValues();
      for (var i = 1; i < raw.length; i++) {
        if (String(raw[i][1] || '').trim() !== item.refId) continue;
        var vivo = _comNum(raw[i][3]);
        // El crédito nació con item.monto. Si hoy vale menos, ya se consumió algo.
        if (Math.abs(vivo - item.monto) > 0.01)
          return { tocado: true, rowNum: i + 1,
                   motivo: 'El crédito de ' + item.beneficiario + ' ya se usó (nació en $' + item.monto.toLocaleString('es-MX') + ' y hoy quedan $' + vivo.toLocaleString('es-MX') + '). Regenerarlo le devolvería un dinero que ya se aplicó.' };
        return { tocado: false, rowNum: i + 1 };
      }
      // Ya no está: alguien lo borró a mano. No lo damos por bueno.
      return { tocado: true, rowNum: -1, motivo: 'No encuentro el crédito ' + item.refId + ' de ' + item.beneficiario + ' en Creditos_Favor (¿se borró a mano?). Revísalo antes de regenerar.' };
    }
    if (item.via === 'efectivo') {
      var ssE = SpreadsheetApp.openById(EGRESOS_SS_2026);
      var shE = ssE.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
      if (!shE) return { tocado: true, rowNum: -1, motivo: 'No se encontró la hoja de Egresos.' };
      var rawE = shE.getDataRange().getValues();
      for (var j = 1; j < rawE.length; j++) {
        if (String(rawE[j][0] || '').trim() !== String(item.refId)) continue;
        var pagado = rawE[j][13] === true || String(rawE[j][13]).toUpperCase() === 'TRUE';
        var fecha = rawE[j][1];
        if (pagado || fecha)
          return { tocado: true, rowNum: j + 1,
                   motivo: 'La comisión en efectivo de ' + item.beneficiario + ' ($' + item.monto.toLocaleString('es-MX') + ') YA SE PAGÓ. Borrarla eliminaría el registro de un pago real.' };
        return { tocado: false, rowNum: j + 1 };
      }
      return { tocado: true, rowNum: -1, motivo: 'No encuentro la cuenta por pagar ' + item.refId + ' de ' + item.beneficiario + ' en Egresos (¿se borró a mano?). Revísala antes de regenerar.' };
    }
    return { tocado: true, rowNum: -1, motivo: 'Vía desconocida: ' + item.via };
  } catch (ex) { return { tocado: true, rowNum: -1, motivo: 'No se pudo revisar ' + item.refId + ': ' + ex.message }; }
}

/* ═════════════════════════ GENERAR (el que escribe) ═════════════════════════ */
function generarComisiones(body) {
  try {
    body = body || {};
    var dep = _comGuardDeps(); if (dep) return { ok: false, error: dep, version: COMISIONES_VER };

    // PERMISO ANTES QUE NADA: ni una lectura de más si no tiene derecho a escribir.
    // Al ser un permiso nuevo, ningún rol lo trae todavía → solo admin/director
    // (que _tokenHasPermission deja pasar). Falla del lado seguro a propósito.
    if (typeof _tokenHasPermission !== 'function')
      return { ok: false, error: 'Actualiza finance.gs en Apps Script y redespliega (falta el verificador de permisos).', version: COMISIONES_VER };
    if (!_tokenHasPermission(body.token, COMISIONES_PERM))
      return { ok: false, error: 'No tienes permiso para generar comisiones (' + COMISIONES_PERM + '). Pídeselo al administrador o al director.', version: COMISIONES_VER };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado, reintenta en un momento.', version: COMISIONES_VER };
    try {
      // Recalcular AQUÍ DENTRO: nunca confiar en los montos que mande el cliente.
      // Lo que se aprobó en pantalla se vuelve a calcular contra la hoja antes de escribir.
      var calc = calcularComisiones({ anio: body.anio, mes: body.mes, reglaId: body.reglaId });
      if (!calc.ok) return calc;
      if (calc.bloqueo) return { ok: false, error: calc.bloqueo, version: COMISIONES_VER };
      if (!calc.detalle.length)
        return { ok: false, error: 'No hay nada que generar: el grupo hizo ' + calc.conteo + ' procedimiento(s) y le corresponde ' + calc.pct + '%.', version: COMISIONES_VER, calculo: calc };

      var mesKey = calc.mes, reglaId = calc.regla.id;
      var ya = _comYaGenerado(mesKey, reglaId);

      // ── SOLO REVERTIR (deshacer una generación hecha por error) ──────────
      // Misma seguridad que regenerar: se revisa TODO antes de tocar NADA; si una
      // pieza ya se cobró/usó, NO se revierte nada (para no borrar dinero real).
      if (body.soloRevertir) {
        if (!ya.length) return { ok: false, error: 'No hay comisiones generadas de ' + mesKey + ' (regla ' + reglaId + ') para revertir.', version: COMISIONES_VER };
        var estR = ya.map(function (i) { return { item: i, est: _comEstadoRef(i) }; });
        var tocR = estR.filter(function (e) { return e.est.tocado; });
        if (tocR.length)
          return { ok: false, version: COMISIONES_VER, noSePuedeRevertir: true,
                   error: 'No se puede revertir ' + mesKey + ': parte de lo generado ya se cobró/usó.\n· ' +
                          tocR.map(function (t) { return t.est.motivo; }).join('\n· ') + '\nNo se modificó nada.',
                   motivos: tocR.map(function (t) { return t.est.motivo; }) };
        var revR = [];
        estR.forEach(function (e) {
          var i = e.item;
          if (i.via === 'nota_credito') { _cobRegistrarCreditoFavor(i.refId, i.beneficiario, 0, null); }
          else if (i.via === 'efectivo' && e.est.rowNum > 1) {
            var ssDr = SpreadsheetApp.openById(EGRESOS_SS_2026);
            var shDr = ssDr.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
            shDr.deleteRow(e.est.rowNum);
            try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (_c) {}
          }
          revR.push({ beneficiario: i.beneficiario, via: i.via, monto: i.monto, refId: i.refId });
        });
        var shCr = _comEnsureControl();
        ya.map(function (i) { return i.rowNum; }).sort(function (a, b) { return b - a; })
          .forEach(function (rn) { shCr.getRange(rn, 7).setValue('revertida'); });
        try { logAudit(body.usuario || 'sistema', 'Comisiones', 'Revertir (deshacer)', mesKey + '/' + reglaId, '', '', revR.length + ' movimiento(s)'); } catch (e) {}
        return { ok: true, revertido: true, version: COMISIONES_VER, mes: mesKey, reglaId: reglaId, revertidos: revR };
      }

      // ── IDEMPOTENCIA ────────────────────────────────────────────────────
      if (ya.length && !body.regenerar) {
        return { ok: false, yaGenerado: true, version: COMISIONES_VER,
                 error: 'Las comisiones de ' + mesKey + ' (regla ' + reglaId + ') YA SE GENERARON el ' +
                        (ya[0].timestamp ? _cobStr(_cobD(ya[0].timestamp)) : '—') + ' por ' + (ya[0].usuario || '—') + ': ' +
                        ya.map(function (i) { return i.beneficiario + ' $' + i.monto.toLocaleString('es-MX') + ' (' + i.via + ')'; }).join(', ') +
                        '. No se generó nada de nuevo. Si necesitas rehacerlas, usa «Regenerar» — solo funciona si nadie las ha cobrado todavía.',
                 items: ya };
      }
      var revertidos = [];
      if (ya.length && body.regenerar) {
        // 1) Revisar TODO primero. Si UNA sola pieza ya se cobró, no se toca NADA.
        var estados = ya.map(function (i) { return { item: i, est: _comEstadoRef(i) }; });
        var tocados = estados.filter(function (e) { return e.est.tocado; });
        if (tocados.length)
          return { ok: false, version: COMISIONES_VER, noSePuedeRegenerar: true,
                   error: 'No se puede regenerar ' + mesKey + ': parte de lo generado ya se usó y regenerarlo duplicaría dinero.\n· ' +
                          tocados.map(function (t) { return t.est.motivo; }).join('\n· ') +
                          '\nNo se modificó nada. Corrige a mano lo que sobre, o genera el ajuste como un movimiento aparte.',
                   motivos: tocados.map(function (t) { return t.est.motivo; }) };
        // 2) Nadie lo tocó → revertir.
        estados.forEach(function (e) {
          var i = e.item;
          if (i.via === 'nota_credito') {
            // monto 0 → el crédito sale de "disponible" sin borrar el renglón (deja rastro).
            _cobRegistrarCreditoFavor(i.refId, i.beneficiario, 0, null);
          } else if (i.via === 'efectivo' && e.est.rowNum > 1) {
            var ssD = SpreadsheetApp.openById(EGRESOS_SS_2026);
            var shD = ssD.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
            shD.deleteRow(e.est.rowNum);
            try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (_c) {}
          }
          revertidos.push({ beneficiario: i.beneficiario, via: i.via, monto: i.monto, refId: i.refId });
        });
        // 3) Marcar el control (de abajo hacia arriba: borrar filas recorre índices).
        var shC = _comEnsureControl();
        ya.map(function (i) { return i.rowNum; }).sort(function (a, b) { return b - a; })
          .forEach(function (rn) { shC.getRange(rn, 7).setValue('revertida'); });
        try { logAudit(body.usuario || 'sistema', 'Comisiones', 'Revertir', mesKey + '/' + reglaId, '', '', revertidos.length + ' movimiento(s)'); } catch (e) {}
      }

      // ── ESCRIBIR ────────────────────────────────────────────────────────
      // Se escribe UN movimiento POR BENEFICIARIO (calc.totales), no uno por línea de
      // detalle: el coordinador cobra un solo pago al mes por los procedimientos de todos
      // los médicos, no un pago suelto por cada uno.
      var creados = [], errores = [];
      calc.totales.forEach(function (T) {
        var desglose = T.lineas.map(function (l) { return l.medicoNombre + ' $' + l.base.toLocaleString('es-MX') + ' → $' + l.monto.toLocaleString('es-MX'); }).join('; ');
        var pctTxt = T.lineas.length ? (T.lineas[0].pct + '% × ' + T.lineas[0].parte + '% = ' + T.lineas[0].efectivoPct + '%') : '';
        var concepto = 'Comisión ' + (calc.regla.nombre || reglaId) + ' ' + mesKey +
                       ' · escalón ' + pctTxt + ' · ' + desglose;
        if (T.via === 'efectivo') {
          var venc = _comFinDeMes(T.mesAplica);
          var res = saveCxP({ mes: T.mesAplica, proveedor: T.beneficiario,
                              contable: 'Gasto', tipo: 'Variable', subtipo: 'Comisiones',
                              concepto: concepto, monto: T.monto, vencimiento: venc,
                              notas: 'Generada por el módulo de Comisiones (' + reglaId + '/' + mesKey + '). Se paga en efectivo.',
                              usuario: body.usuario || 'sistema' });
          if (!res || !res.ok) { errores.push('CxP de ' + T.beneficiario + ': ' + ((res && res.error) || 'error desconocido')); return; }
          creados.push({ beneficiario: T.beneficiario, via: 'efectivo', monto: T.monto, refId: String(res.id), vencimiento: venc });
        } else if (T.via === 'nota_credito') {
          var op = _comOpCredito(reglaId, mesKey, T.beneficiarioId);
          _cobRegistrarCreditoFavor(op, T.beneficiario, T.monto, _comFinDeMes(mesKey));
          // _cobRegistrarCreditoFavor se traga TODOS sus errores (try/catch vacío) y no
          // devuelve nada: si fallara, escribiríamos en el control un crédito que no existe.
          // Se verifica leyendo de vuelta.
          var chk = _comEstadoRef({ via: 'nota_credito', refId: op, monto: T.monto, beneficiario: T.beneficiario });
          if (chk.tocado || chk.rowNum < 0) { errores.push('Crédito de ' + T.beneficiario + ': no quedó registrado en Creditos_Favor.'); return; }
          creados.push({ beneficiario: T.beneficiario, via: 'nota_credito', monto: T.monto, refId: op, mesAplica: T.mesAplica });
        }
      });

      // Registrar en el control SOLO lo que de verdad se escribió.
      if (creados.length) {
        var shCtl = _comEnsureControl();
        var stamp = new Date();
        shCtl.getRange(shCtl.getLastRow() + 1, 1, creados.length, 9).setValues(creados.map(function (c) {
          return [mesKey, reglaId, c.beneficiario, c.via, c.monto, c.refId, 'activa', body.usuario || 'sistema', stamp];
        }));
      }
      try {
        logAudit(body.usuario || 'sistema', 'Comisiones', body.regenerar ? 'Regenerar' : 'Generar',
                 mesKey + '/' + reglaId, 'Escalón', calc.conteo + ' proc.',
                 calc.pct + '% · ' + creados.length + ' movimiento(s) · $' + calc.totalGeneral);
      } catch (e) {}

      return { ok: errores.length === 0, version: COMISIONES_VER,
               mes: mesKey, reglaId: reglaId, conteo: calc.conteo, pct: calc.pct,
               creados: creados, revertidos: revertidos, errores: errores,
               total: _comRedondea(creados.reduce(function (s, c) { return s + c.monto; }, 0)),
               error: errores.length ? ('Se generaron ' + creados.length + ' movimiento(s), pero fallaron otros:\n· ' + errores.join('\n· ')) : '' };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═════════════════════════ SETUP ═════════════════════════ */
function setupComisionesConfig() {
  try {
    var cfg = _comCfg();
    PropertiesService.getScriptProperties().setProperty(COMISIONES_CFG_KEY, JSON.stringify(cfg));
    _comEnsureControl();
    return { ok: true, version: COMISIONES_VER, reglas: (cfg.reglas || []).length, tab: COMISIONES_TAB,
             deploy: _comGuardDeps() || '' };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}
