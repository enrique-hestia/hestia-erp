/**
 * cobranza.gs — Módulo de Cuentas por Cobrar (Cobranza) para Hestia Fertility.
 *
 * DOS motores independientes (no revuelven datos):
 *   Motor A — Saldos por cobrar (adeudos por pago incompleto).
 *             Cargo = venta en BD_Ingresos (TotalPagar).
 *             Abono = pago recibido, registrado en hoja Abonos_Cobrar.
 *             Saldo = ΣTotalPagar − ΣPagado(col) − Σabonos.
 *             Los históricos con Pagado en blanco se asumen PAGADOS (saldo 0);
 *             el AR nace cuando se captura un pago parcial o se carga un saldo inicial.
 *             Registro manual extra en hoja Cuentas_Cobrar (saldos iniciales / cargos a crédito).
 *
 *   Motor B — Suscripciones de Crío (recurrente).
 *             Universo = pacientes con inventario en Inventario Crío (LAB).
 *             Ancla = Fecha Crío. Año 1 GRATIS solo para NO-externos (Hestia).
 *             Anual: cada pago cubre 1 año (paga 1/ene → cubre a 31/dic → próximo cobro 1/ene sig).
 *             Mensual: se debe el día 1 de cada mes (acumula atrasos).
 *
 * Requiere globals ya definidos en el proyecto:
 *   INGRESOS_SS_ID, BD_INGRESOS_TAB, LAB_SS_ID, findSheet (lab.gs),
 *   _privVer / _privPaciente (privacidad.gs).
 *
 * Funciones públicas (se cablean en core.gs GET y finance.gs POST):
 *   readCuentasPorCobrar(body)   — Motor A (lista + totales + aging)
 *   readSuscripcionesCrio(body)  — Motor B (lista + totales + estatus)
 *   readEstadoCobranza(paciente) — estado de cuenta de cobranza por paciente
 *   registrarAbono(body)         — registra un abono a un saldo o suscripción
 *   cargarSaldoInicial(body)     — carga un saldo inicial / cargo a crédito
 *   setupCobranzaConfig()        — inicializa config + hojas
 */

var COBRANZA_CFG_KEY  = 'COBRANZA_CONFIG';
var COBRANZA_ABONOS   = 'Abonos_Cobrar';
var COBRANZA_CARGOS   = 'Cuentas_Cobrar';
var COBRANZA_SUS      = 'Suscripciones_Crio';
var COBRANZA_VER      = 'cobranza-2026.07.10m';

/* ───────────────────────── Config ───────────────────────── */
function _cobCfg() {
  var def = {
    tarifaAnual:          5700,     // MXN al año
    tarifaMensual:        475,      // MXN al mes
    anioGratisSoloHestia: true,     // año 1 gratis SOLO para no-externos
    externosPaganDesdeInicio: true, // externos: sin año gratis
    planDefault:          'anual',  // cuando el paciente tiene crío pero nunca eligió plan
    diasPorVencer:        30,       // ventana para "Por vencer"
    maxMesesAtraso:       36        // tope de meses de atraso a cobrar (evita montos absurdos)
  };
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(COBRANZA_CFG_KEY);
    if (raw) {
      var o = JSON.parse(raw);
      for (var k in def) { if (!(k in o)) o[k] = def[k]; }
      return o;
    }
  } catch (e) {}
  return def;
}

function setupCobranzaConfig() {
  var cfg = _cobCfg();
  PropertiesService.getScriptProperties().setProperty(COBRANZA_CFG_KEY, JSON.stringify(cfg));
  _cobEnsureAbonos();
  _cobEnsureCargos();
  _susEnsureSheet();
  return { ok: true, version: COBRANZA_VER, config: cfg };
}

/* ── Configuración por paciente (suscripción crío) ──
 * Por DEFAULT: Hestia (año 1 gratis) + plan ANUAL. Flags manuales por paciente:
 *   externo → paga desde el año 1 (sin año gratis)
 *   mensual → inscrito a pago mensual (registra vencimientos mensuales; el default
 *             es anual — el plan mensual es opt-in, no se deriva del producto)
 *   autopay → pago automático (solo informativo; se confirma el cobro del próximo año)
 * No se usa la columna EXTERNO del Inventario Crío (no es clasificación de cobro). */
var COBRANZA_PAC_KEY = 'COBRANZA_PAC_CFG';
function _cobPacCfg() {
  try { var raw = PropertiesService.getScriptProperties().getProperty(COBRANZA_PAC_KEY); if (raw) return JSON.parse(raw) || {}; } catch (e) {}
  return {};
}
function _cobSetPacFlag(paciente, flag, value) {
  var key = _cobKeyNom(paciente); if (!key) return;
  var cfg = _cobPacCfg();
  if (!cfg[key]) cfg[key] = {};
  cfg[key][flag] = !!value;
  cfg[key]._nombre = String(paciente).trim();
  if (!cfg[key].externo && !cfg[key].mensual && !cfg[key].autopay) delete cfg[key];
  PropertiesService.getScriptProperties().setProperty(COBRANZA_PAC_KEY, JSON.stringify(cfg));
}
function cobPacienteFlag(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_egresos'))
      return { ok: false, error: 'Sin permiso (requiere editar cuentas por pagar).' };
    var flag = String(body.flag || '').toLowerCase();
    if (['externo', 'mensual', 'autopay'].indexOf(flag) < 0) return { ok: false, error: 'Flag inválido' };
    if (!body.paciente) return { ok: false, error: 'Paciente inválido' };
    var val = (body.value === true || body.value === 'true');
    _cobSetPacFlag(body.paciente, flag, val);
    return { ok: true, version: COBRANZA_VER, flag: flag, value: val, paciente: String(body.paciente).trim() };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_VER }; }
}
// Back-compat: el botón "Tipo" del frontend manda action=cobExterno.
function cobExterno(body) { body.flag = 'externo'; body.value = (body.externo === true || body.externo === 'true'); return cobPacienteFlag(body); }

/* ───────────────────────── Helpers ───────────────────────── */
function _cobNum(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v || '').replace(/[$,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
function _cobIsBlank(v) { return v === '' || v === null || v === undefined; }

function _cobD(v) {  // → Date (sin hora) o null
  if (!v && v !== 0) return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function _cobStr(d) {  // Date → yyyy-MM-dd
  if (!d) return '';
  if (!(d instanceof Date)) { d = _cobD(d); if (!d) return ''; }
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function _cobToday() { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }
function _cobAddYears(d, n)  { return new Date(d.getFullYear() + n, d.getMonth(), d.getDate()); }
function _cobAddMonths(d, n) { return new Date(d.getFullYear(), d.getMonth() + n, d.getDate()); }
function _cobStartMonth(d)   { return new Date(d.getFullYear(), d.getMonth(), 1); }
function _cobDaysDiff(a, b)  { return Math.round((b.getTime() - a.getTime()) / 86400000); }
function _cobMonthsDiff(a, b) { // meses completos de a → b (por día 1 de mes)
  return (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());
}
function _cobLower(v) { return String(v || '').trim().toLowerCase(); }
function _cobKeyNom(v) { // clave de paciente normalizada (sin acentos/espacios extra)
  return _cobLower(v).replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
    .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u').replace(/ñ/g, 'n').replace(/\s+/g, ' ').trim();
}
function _cobMasked() { return (typeof _privVer === 'function' && !_privVer()); }

function _cobPlan(nombre) {  // detecta plan por nombre de producto
  var s = _cobLower(nombre);
  if (/mensual/.test(s)) return { plan: 'mensual', meses: 1 };
  if (/anual|12\s*mes|anualidad/.test(s)) return { plan: 'anual', meses: 12 };
  if (/6\s*mes|semestr/.test(s)) return { plan: 'semestral', meses: 6 };
  return null;
}
function _cobEsAlmacenamiento(cat, prod) {
  var s = _cobLower(cat) + ' ' + _cobLower(prod);
  return /almacen|criopreserv|congelamiento|mantenimiento congel|anualidad|suscrip/.test(s);
}

/* ───────────────── Segmentación (Motor A) ───────────────── */
function _cobSegmento(cat, prod) {
  var s = _cobLower(cat) + ' ' + _cobLower(prod);
  if (/reprovida|agencia/.test(s)) return 'Agencias';
  if (/grupo\s*medico|grupomedico|externo|medico/.test(s)) return 'Médicos';
  return 'Pacientes';
}

/* ───────────────── Hojas de apoyo ───────────────── */
function _cobEnsureAbonos() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_ABONOS);
  if (!sh) {
    sh = ss.insertSheet(COBRANZA_ABONOS);
    sh.appendRow(['Fecha', 'OP', 'Paciente', 'Monto', 'FormaPago', 'Banco', 'Tipo', 'Nota', 'Usuario', 'Timestamp']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function _cobEnsureCargos() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_CARGOS);
  if (!sh) {
    sh = ss.insertSheet(COBRANZA_CARGOS);
    sh.appendRow(['Fecha', 'OP', 'Paciente', 'Categoria', 'Concepto', 'MontoCargo', 'Estatus', 'Nota', 'Usuario', 'Timestamp', 'Items']);
    sh.setFrozenRows(1);
  }
  return sh;
}
// Índice (1-based) de la columna 'Items' (JSON de partidas); la crea si no existe.
function _cobCargosItemsCol(sh) {
  var lc = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var i = hdr.indexOf('items');
  if (i > -1) return i + 1;
  var col = lc + 1; sh.getRange(1, col).setValue('Items'); return col;
}
// Normaliza las partidas [{producto,cantidad,precio,total}]; total = precio*cant.
function _cobParseItems(x) {
  if (!x) return [];
  var arr = x;
  if (typeof x === 'string') { var s = x.trim(); if (!s) return []; try { arr = JSON.parse(s); } catch (e) { return []; } }
  if (!arr || !arr.length) return [];
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var it = arr[i] || {};
    var c = _cobNum(it.cantidad) || 1;
    var p = _cobNum(it.precio);
    var t = (it.total != null && it.total !== '') ? _cobNum(it.total) : (p * c);
    var pac = String(it.pac || it.paciente || '').trim(); // nombre del paciente real (agencias/externos)
    if (String(it.producto || '').trim() || t > 0) out.push({ producto: String(it.producto || '').trim(), cantidad: c, precio: p, total: t, saldo: 0, pac: pac });
  }
  return out;
}
function _cobItemsMonto(items) { var s = 0; for (var i = 0; i < items.length; i++) s += items[i].total; return s; }
function _cobReadAbonos() {
  var out = { byOp: {}, byPac: {}, susByPac: {}, rows: [] };
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_ABONOS);
  if (!sh) return out;
  var raw = sh.getDataRange().getValues();
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var op = String(r[1] || '').trim();
    var pac = String(r[2] || '').trim();
    var monto = _cobNum(r[3]);
    var tipo = _cobLower(r[6]);
    var fecha = _cobD(r[0]);
    if (monto <= 0 && tipo !== 'suscripcion') continue;
    var rec = { fecha: _cobStr(fecha), op: op, paciente: pac, monto: monto, formaPago: String(r[4] || ''), banco: String(r[5] || ''), tipo: tipo, nota: String(r[7] || '') };
    out.rows.push(rec);
    if (tipo === 'suscripcion') {
      var kp = _cobKeyNom(pac);
      if (!out.susByPac[kp]) out.susByPac[kp] = [];
      out.susByPac[kp].push(rec);
    } else {
      if (op) out.byOp[op] = (out.byOp[op] || 0) + monto;
      var kp2 = _cobKeyNom(pac);
      if (!out.byPac[kp2]) out.byPac[kp2] = 0;
      out.byPac[kp2] += monto;
    }
  }
  return out;
}
function _cobReadCargos() {  // registro manual de saldos iniciales / cargos a crédito
  var out = { byOp: {}, rows: [] };
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_CARGOS);
  if (!sh) return out;
  var raw = sh.getDataRange().getValues();
  var hdr0 = (raw[0] || []).map(function (x) { return String(x).trim().toLowerCase(); });
  var iItems = hdr0.indexOf('items');
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var op = String(r[1] || '').trim();
    var monto = _cobNum(r[5]);
    if (monto <= 0) continue;
    var rec = {
      rowNum: i + 1, fecha: _cobStr(_cobD(r[0])), op: op, paciente: String(r[2] || ''),
      categoria: String(r[3] || ''), concepto: String(r[4] || ''),
      monto: monto, estatus: String(r[6] || ''), nota: String(r[7] || ''),
      items: (iItems > -1) ? _cobParseItems(r[iItems]) : []
    };
    out.rows.push(rec);
    if (op) out.byOp[op] = rec;
  }
  return out;
}
// Registra automáticamente una cuenta por cobrar cuando un ingreso se captura con
// pago parcial. La llama saveIngreso (finance.gs). Nunca debe tumbar la venta.
function _cobRegistrarSaldoIngreso(op, paciente, categoria, monto, fecha) {
  try {
    if (!(_cobNum(monto) > 0)) return;
    var sh = _cobEnsureCargos();
    sh.appendRow([
      fecha ? _cobStr(_cobD(fecha)) : _cobStr(_cobToday()),
      String(op || '').trim(), String(paciente || '').trim(), String(categoria || ''),
      'Saldo pendiente de OP ' + op, _cobNum(monto), 'Pendiente', 'auto-ingreso', '', new Date()
    ]);
  } catch (e) {}
}
// Editar / borrar una cuenta por cobrar. Permisos AMARRADOS a los de Cuentas por
// Pagar (editar_egresos / borrar_egresos) para no crear permisos nuevos.
function editarCuentaCobrar(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_egresos'))
      return { ok: false, error: 'Sin permiso para editar cuentas por cobrar (requiere "Editar partidas de egreso" / cuentas por pagar).' };
    var rn = parseInt(body.rowNum); if (!rn || rn < 2) return { ok: false, error: 'Fila inválida' };
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(COBRANZA_CARGOS);
    if (!sh || rn > sh.getLastRow()) return { ok: false, error: 'Cuenta por cobrar no encontrada' };
    // Cuentas_Cobrar: Fecha(1),OP(2),Paciente(3),Categoria(4),Concepto(5),MontoCargo(6),Estatus(7),Nota(8),...,Items
    // Partidas: si vienen items, se guardan y el MONTO = suma de partidas.
    var _itemsEdit = (body.items != null) ? _cobParseItems(body.items) : null;
    if (_itemsEdit != null) {
      var iItemsE = _cobCargosItemsCol(sh);
      sh.getRange(rn, iItemsE).setValue(_itemsEdit.length ? JSON.stringify(_itemsEdit) : '');
      if (_itemsEdit.length) sh.getRange(rn, 6).setValue(_cobItemsMonto(_itemsEdit));
    }
    if (!(_itemsEdit && _itemsEdit.length) && body.monto != null && String(body.monto) !== '') sh.getRange(rn, 6).setValue(_cobNum(body.monto));
    if (body.concepto != null) sh.getRange(rn, 5).setValue(String(body.concepto));
    if (body.paciente != null && String(body.paciente) !== '') sh.getRange(rn, 3).setValue(String(body.paciente));
    if (body.estatus != null && String(body.estatus) !== '') sh.getRange(rn, 7).setValue(String(body.estatus));
    return { ok: true, version: COBRANZA_VER, rowNum: rn };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_VER }; }
}
function borrarCuentaCobrar(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'borrar_egresos'))
      return { ok: false, error: 'Sin permiso para borrar cuentas por cobrar (requiere "Eliminar partidas de egreso" / cuentas por pagar).' };
    var rn = parseInt(body.rowNum); if (!rn || rn < 2) return { ok: false, error: 'Fila inválida' };
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(COBRANZA_CARGOS);
    if (!sh || rn > sh.getLastRow()) return { ok: false, error: 'Cuenta por cobrar no encontrada' };
    sh.deleteRow(rn);
    return { ok: true, version: COBRANZA_VER, rowNum: rn };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_VER }; }
}
// Ajusta el Pagado de una OP en BD_Ingresos (para saldos que vienen de un pago
// parcial). saldar=true → marca pagado completo (borra el adeudo). Si no, fija el
// monto pagado. Permiso amarrado a editar cuentas por pagar.
function ajustarPagadoIngreso(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_egresos'))
      return { ok: false, error: 'Sin permiso (requiere "Editar partidas de egreso" / cuentas por pagar).' };
    var op = String(body.op || '').trim(); if (!op) return { ok: false, error: 'OP inválida' };
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(BD_INGRESOS_TAB); if (!sh) return { ok: false, error: 'No hay BD_Ingresos' };
    var data = sh.getDataRange().getValues();
    var H = data[0].map(function (x) { return _cobLower(x); });
    var hc = function () { for (var a = 0; a < arguments.length; a++) { var k = H.indexOf(arguments[a]); if (k > -1) return k; } return -1; };
    var iOp = hc('op'); if (iOp < 0) iOp = 0;
    var iTotal = hc('totalpagar', 'total a pagar', 'total'), iPag = hc('pagado');
    if (iPag < 0) return { ok: false, error: 'No se encontró la columna Pagado' };
    var rows = [], totalOP = 0;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iOp] || '').trim() === op) { var t = _cobNum(data[r][iTotal]); rows.push({ rowNum: r + 1, total: t }); totalOP += t; }
    }
    if (!rows.length) return { ok: false, error: 'OP no encontrada' };
    var nuevoPagado = (body.saldar === true || body.saldar === 'true') ? totalOP : _cobNum(body.pagado);
    if (nuevoPagado > totalOP) nuevoPagado = totalOP; if (nuevoPagado < 0) nuevoPagado = 0;
    var rem = nuevoPagado;
    for (var k = 0; k < rows.length; k++) { var apply = Math.min(rem, rows[k].total); if (apply < 0) apply = 0; sh.getRange(rows[k].rowNum, iPag + 1).setValue(apply); rem -= apply; }
    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (e) {}
    return { ok: true, version: COBRANZA_VER, op: op, pagado: nuevoPagado, total: totalOP, saldo: Math.max(0, totalOP - nuevoPagado) };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_VER }; }
}

/* ═════════════════ MOTOR A — Saldos por cobrar ═════════════════
 * Cuentas por Cobrar lee SOLO el registro explícito (hoja Cuentas_Cobrar). NO se
 * deriva de la columna Pagado de BD_Ingresos: ese Pagado es de años pasados con
 * precios más bajos, mientras el TotalPagar trae el precio actual (subió), así que
 * la resta daba "saldos" que NO son deuda. Un adeudo nace solo cuando: (a) se
 * captura un ingreso con pago parcial (saveIngreso lo registra), (b) se carga un
 * saldo inicial, o (c) se registra un cargo a crédito. Cada OP trae su desglose.
 */
function _cobBuildSaldos() {
  var cargos = _cobReadCargos();
  var abonos = _cobReadAbonos();
  var today = _cobToday();
  // Desglose (items) SOLO para las OPs que están en el registro — para ver qué es
  // cada OP sin ir a la venta. No se derivan saldos de BD_Ingresos.
  var opsNeeded = {}, pacsNeeded = {};
  for (var q = 0; q < cargos.rows.length; q++) {
    if (cargos.rows[q].op) opsNeeded[cargos.rows[q].op] = true;
    if (cargos.rows[q].paciente) pacsNeeded[_cobKeyNom(cargos.rows[q].paciente)] = true;
  }
  var itemsByOp = _cobMasked() ? {} : _cobItemsForOps(opsNeeded);
  var itemsByPac = _cobMasked() ? {} : _cobItemsForPacientes(pacsNeeded);
  var out = [];
  for (var ci = 0; ci < cargos.rows.length; ci++) {
    var cg = cargos.rows[ci];
    var st = _cobLower(cg.estatus);
    if (st === 'cancelado' || st === 'pagado') continue;
    var abo = cg.op ? (abonos.byOp[cg.op] || 0) : 0;
    var saldo = cg.monto - abo;
    if (saldo <= 0.01) continue;
    var fd = _cobD(cg.fecha) || today;
    var dias = _cobDaysDiff(fd, today); if (dias < 0) dias = 0;
    // Detalle (redespliegue): si el cargo tiene OP se usa el desglose de esa OP; si
    // no (ej. REPROVIDA con saldo inicial), se jala TODO lo vendido a ese paciente.
    var det = [];
    if (!_cobMasked()) {
      if (cg.items && cg.items.length) det = cg.items;          // partidas capturadas a mano
      else if (cg.op && itemsByOp[cg.op]) det = itemsByOp[cg.op]; // por OP
      else det = itemsByPac[_cobKeyNom(cg.paciente)] || [];       // por paciente (ej. REPROVIDA)
    }
    out.push({
      origen: 'registro', rowNum: cg.rowNum, op: cg.op || '—',
      paciente: _cobMasked() ? (cg.op ? _privPaciente(cg.op) : 'Paciente') : cg.paciente,
      segmento: _cobSegmento(cg.categoria, cg.concepto),
      categoria: cg.categoria, concepto: cg.concepto || cg.estatus || 'Saldo',
      fecha: cg.fecha, total: cg.monto, abonado: abo, saldo: saldo,
      dias: dias, bucket: _cobBucket(dias),
      items: det
    });
  }
  out.sort(function (a, b) { return b.saldo - a.saldo; });
  return { ops: out };
}
// Trae las líneas de una lista de OPs desde BD_Ingresos (solo para el desglose).
function _cobItemsForOps(opsNeeded) {
  var res = {};
  if (!opsNeeded || !Object.keys(opsNeeded).length) return res;
  var sh = SpreadsheetApp.openById(INGRESOS_SS_ID).getSheetByName(BD_INGRESOS_TAB);
  if (!sh) return res;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return res;
  var H = data[0].map(function (x) { return _cobLower(x); });
  var hc = function () { for (var a = 0; a < arguments.length; a++) { var k = H.indexOf(arguments[a]); if (k > -1) return k; } return -1; };
  var iOp = hc('op'); if (iOp < 0) iOp = 0;
  var iProd = hc('producto'), iCant = hc('cantidad'), iTotal = hc('totalpagar', 'total a pagar', 'total');
  for (var r = 1; r < data.length; r++) {
    var op = String(data[r][iOp] || '').trim(); if (!opsNeeded[op]) continue;
    if (!res[op]) res[op] = [];
    if (res[op].length < 30 && String(data[r][iProd] || '').trim()) res[op].push({ producto: String(data[r][iProd] || '').trim(), cantidad: iCant > -1 ? _cobNum(data[r][iCant]) : 1, total: _cobNum(data[r][iTotal]), saldo: 0 });
  }
  return res;
}
// Igual que _cobItemsForOps pero por PACIENTE (para saldos sin OP, ej. REPROVIDA).
function _cobItemsForPacientes(pacsNeeded) {
  var res = {};
  if (!pacsNeeded || !Object.keys(pacsNeeded).length) return res;
  var sh = SpreadsheetApp.openById(INGRESOS_SS_ID).getSheetByName(BD_INGRESOS_TAB);
  if (!sh) return res;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return res;
  var H = data[0].map(function (x) { return _cobLower(x); });
  var hc = function () { for (var a = 0; a < arguments.length; a++) { var k = H.indexOf(arguments[a]); if (k > -1) return k; } return -1; };
  var iPac = hc('paciente'), iOp = hc('op'), iFecha = hc('fecha'), iProd = hc('producto'), iCant = hc('cantidad'),
      iTotal = hc('totalpagar', 'total a pagar', 'total'), iPag = hc('pagado');
  for (var r = 1; r < data.length; r++) {
    var pk = _cobKeyNom(data[r][iPac]); if (!pacsNeeded[pk]) continue;
    if (!res[pk]) res[pk] = [];
    if (res[pk].length < 60 && String(data[r][iProd] || '').trim()) {
      var tot = _cobNum(data[r][iTotal]); var pag = iPag > -1 ? _cobNum(data[r][iPag]) : 0;
      res[pk].push({ producto: String(data[r][iProd] || '').trim(), cantidad: iCant > -1 ? _cobNum(data[r][iCant]) : 1, total: tot, saldo: (pag > 0.01 && pag < tot - 0.01) ? (tot - pag) : 0, fecha: _cobStr(_cobD(data[r][iFecha])), op: String(data[r][iOp] || '').trim() });
    }
  }
  return res;
}
function _cobBucket(dias) {
  if (dias <= 30) return '0-30';
  if (dias <= 60) return '31-60';
  if (dias <= 90) return '61-90';
  return '+90';
}

function readCuentasPorCobrar(body) {
  try {
    var built = _cobBuildSaldos();
    if (built.error) return { ok: false, error: built.error, version: COBRANZA_VER };
    var ops = built.ops;
    var totSaldo = 0, totFacturado = 0, totAbonado = 0;
    var porSeg = { Pacientes: 0, 'Médicos': 0, Agencias: 0 };
    var aging = { '0-30': 0, '31-60': 0, '61-90': 0, '+90': 0 };
    for (var i = 0; i < ops.length; i++) {
      var o = ops[i];
      totSaldo += o.saldo; totFacturado += o.total; totAbonado += o.abonado;
      porSeg[o.segmento] = (porSeg[o.segmento] || 0) + o.saldo;
      aging[o.bucket] += o.saldo;
    }
    return {
      ok: true, version: COBRANZA_VER, view: 'cuentas-cobrar',
      ops: ops, numDeudores: _cobCountDeudores(ops),
      totalSaldo: totSaldo, totalFacturado: totFacturado, totalAbonado: totAbonado,
      porSegmento: porSeg, aging: aging, masked: _cobMasked()
    };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}
function _cobCountDeudores(ops) {
  var s = {}; for (var i = 0; i < ops.length; i++) s[ops[i].paciente] = 1; return Object.keys(s).length;
}

/* ═════════════════ MOTOR B — Suscripciones Crío ═════════════════ */
function _cobReadCrio() {
  // → { pacientes: [{nombre, key, crioInicio(Date), oov, emb, externo}], error? }
  var out = { pacientes: [] };
  var ssLab;
  try { ssLab = SpreadsheetApp.openById(LAB_SS_ID); }
  catch (e) { out.error = 'No se pudo abrir LAB'; return out; }
  var sh = (typeof findSheet === 'function')
    ? (findSheet(ssLab, 'Inventario Crío') || findSheet(ssLab, 'Inventario Crio'))
    : (ssLab.getSheetByName('Inventario Crío') || ssLab.getSheetByName('Inventario Crio'));
  if (!sh) { out.error = 'No se encontró Inventario Crío'; return out; }
  var raw = sh.getDataRange().getValues();
  if (raw.length < 2) return out;
  var H = raw[0].map(function (x) { return _cobLower(x); });
  function hc() { for (var a = 0; a < arguments.length; a++) { for (var j = 0; j < H.length; j++) { if (H[j].indexOf(arguments[a]) > -1) return j; } } return -1; }
  var iNom = hc('nombre', 'paciente'); if (iNom < 0) iNom = 0;
  var iFec = hc('fecha crío', 'fecha crio', 'fecha'); if (iFec < 0) iFec = 1;
  var iOov = hc('oov', 'ovoc'); if (iOov < 0) iOov = 2;
  var iEmb = hc('emb'); if (iEmb < 0) iEmb = 3;
  var iInc = hc('inc emb', 'inc. emb', 'incemb', 'inc emb.', 'inc');  // embriones incluidos
  var iExt = hc('externo', 'hestia');

  var map = {};
  for (var r = 1; r < raw.length; r++) {
    var row = raw[r];
    var nom = String(row[iNom] || '').trim();
    if (!nom) continue;
    var key = _cobKeyNom(nom);
    if (!key) continue;
    var f = _cobD(row[iFec]);
    var oov = _cobNum(row[iOov]);
    var emb = _cobNum(row[iEmb]);
    var inc = iInc > -1 ? _cobNum(row[iInc]) : 0;
    var extRaw = iExt > -1 ? row[iExt] : '';
    var esExtLab = _cobEsExterno(extRaw);
    if (!map[key]) map[key] = { nombre: nom, key: key, crioInicio: f, oov: 0, emb: 0, incEmb: 0, externo: false, externoLab: false };
    var m = map[key];
    if (f && (!m.crioInicio || f.getTime() < m.crioInicio.getTime())) m.crioInicio = f;
    m.oov += oov; m.emb += emb; m.incEmb += inc;
    if (iExt > -1 && !_cobIsBlank(extRaw)) { m.externoLab = m.externoLab || esExtLab; }
  }
  // Config de cobro por paciente (default Hestia + anual). SOLO se cobra a quien
  // tenga inventario (OOV o EMB > 0): en ceros no hay nada que almacenar → no se cobra.
  var pacCfg = _cobPacCfg();
  for (var k in map) {
    var mm = map[k];
    if ((mm.oov + mm.emb) <= 0) continue;  // sin inventario → no se cobra almacenamiento
    var pc = pacCfg[mm.key] || {};
    mm.externo = !!pc.externo;
    mm.mensual = !!pc.mensual;
    mm.autopay = !!pc.autopay;
    out.pacientes.push(mm);
  }
  return out;
}
function _cobEsExterno(v) {
  if (_cobIsBlank(v)) return false;
  if (v === true) return true;
  if (v === 1) return true;
  var s = _cobLower(v);
  if (s === 'externo' || s === 'ext' || s === 'true' || s === '1' || s === 'si' || s === 'sí' || s === 'x') return true;
  return false;
}

function _cobStoragePagosByPac() {
  // compras de almacenamiento en BD_Ingresos → por paciente: [{fecha, producto, monto, plan}]
  var res = {};
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(BD_INGRESOS_TAB);
  if (!sh) return res;
  var data = sh.getDataRange().getValues();
  var H = data[0].map(function (x) { return _cobLower(x); });
  function hc() { for (var a = 0; a < arguments.length; a++) { var k = H.indexOf(arguments[a]); if (k > -1) return k; } return -1; }
  var iFecha = hc('fecha'), iPac = hc('paciente'),
      iCat = hc('categoria', 'categoría'), iProd = hc('producto'),
      iTotal = hc('totalpagar', 'total a pagar', 'total'), iCant = hc('cantidad');
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var cat = String(row[iCat] || ''), prod = String(row[iProd] || '');
    if (!_cobEsAlmacenamiento(cat, prod)) continue;
    var pac = String(row[iPac] || '').trim(); if (!pac) continue;
    var key = _cobKeyNom(pac);
    var f = _cobD(row[iFecha]); if (!f) continue;
    var plan = _cobPlan(prod);
    var cant = iCant > -1 ? _cobNum(row[iCant]) : 1; if (cant <= 0) cant = 1;
    var monto = _cobNum(row[iTotal]);
    if (!res[key]) res[key] = [];
    res[key].push({ fecha: f, producto: prod, monto: monto, precioUnit: monto / cant, plan: plan ? plan.plan : null, meses: plan ? plan.meses : null });
  }
  return res;
}

function _cobCalcSuscripcion(pac, pagos, susAbonos, cfg, today) {
  // Determina plan, última fecha de pago, cobertura, estatus y monto que debe.
  var esExterno = !!pac.externo;
  var crio = pac.crioInicio;
  // El plan es por INSCRIPCIÓN (override), NO se deriva del producto: DEFAULT ANUAL.
  // Solo los inscritos a meses (pac.mensual) usan cobro mensual; los vencidos que no
  // están inscritos se cobran anual (no se difiere a meses).
  var plan = pac.mensual ? 'mensual' : 'anual';
  var meses = plan === 'mensual' ? 1 : 12;
  var tarifa = plan === 'mensual' ? cfg.tarifaMensual : cfg.tarifaAnual;
  var lastPago = null, ultProducto = '';
  var todos = (pagos || []).slice();
  (susAbonos || []).forEach(function (a) { var f = _cobD(a.fecha); if (f) todos.push({ fecha: f, producto: 'Abono suscripción', monto: a.monto }); });
  todos.sort(function (a, b) { return a.fecha.getTime() - b.fecha.getTime(); });
  for (var i = 0; i < todos.length; i++) {
    var p = todos[i];
    if (p.producto) ultProducto = p.producto;
    if (!lastPago || p.fecha.getTime() > lastPago.getTime()) lastPago = p.fecha;
  }
  var tienePlan = todos.length > 0 || !!pac.mensual;

  // Pago automático: solo informativo (se cobra solo). Se confirma el próximo año.
  if (pac.autopay) {
    var covUntil = lastPago ? _cobAddYears(lastPago, 1) : null;
    return {
      plan: plan, tarifa: tarifa, esExterno: esExterno, autopay: true,
      crioInicio: _cobStr(crio), billStart: '', ultimoPago: _cobStr(lastPago), ultProducto: ultProducto,
      coberturaHasta: _cobStr(covUntil ? new Date(covUntil.getTime() - 86400000) : null),
      proximoCobro: _cobStr(covUntil), estatus: 'Pago automático', montoDebe: 0, mesesDebe: 0,
      tienePlan: true, oov: pac.oov, emb: pac.emb
    };
  }

  // fecha de inicio de cobro (billStart): con año gratis para Hestia
  var billStart;
  if (!crio) {
    billStart = lastPago ? _cobStartMonth(lastPago) : today;
  } else if (!esExterno && cfg.anioGratisSoloHestia) {
    billStart = _cobAddYears(crio, 1);   // año 1 gratis
  } else {
    billStart = new Date(crio.getTime());
  }

  var estatus, montoDebe = 0, coberturaHasta = null, proximoCobro = null, mesesDebe = 0;
  // Ventana de cobranza PROACTIVA = hasta el fin del mes SIGUIENTE. Lo que se
  // vence dentro de esta ventana se marca "Por vencer" y ya es cobrable, para no
  // esperar a que esté vencido.
  var finVentana = new Date(today.getFullYear(), today.getMonth() + 2, 0);

  // ¿En cortesía (año gratis) y sin haber pagado aún?
  if (crio && !esExterno && cfg.anioGratisSoloHestia && today.getTime() < billStart.getTime() && !lastPago) {
    proximoCobro = billStart;
    if (billStart.getTime() <= finVentana.getTime()) {
      // el año gratis termina este mes o el próximo → ya cobrable (proactivo)
      estatus = 'Por vencer';
      montoDebe = (plan === 'mensual') ? cfg.tarifaMensual : cfg.tarifaAnual;
    } else {
      estatus = 'Cortesía';
    }
  } else if (plan === 'mensual') {
    // primer cobro: día 1 del mes de billStart (o mes siguiente si ya pasó el día)
    var primerCobro = _cobStartMonth(billStart);
    var cobertoHastaMes;  // primer mes AÚN no cubierto
    if (lastPago) {
      cobertoHastaMes = _cobStartMonth(_cobAddMonths(lastPago, 1));  // el pago cubre su mes; debe desde el siguiente
      if (cobertoHastaMes.getTime() < primerCobro.getTime()) cobertoHastaMes = primerCobro;
    } else {
      cobertoHastaMes = primerCobro;
    }
    var mesesTranscurridos = _cobMonthsDiff(cobertoHastaMes, _cobStartMonth(today)) + 1; // incluye el mes en curso
    if (today.getTime() < cobertoHastaMes.getTime()) mesesTranscurridos = 0;
    mesesDebe = Math.max(0, Math.min(mesesTranscurridos, cfg.maxMesesAtraso));
    coberturaHasta = lastPago ? new Date(_cobAddMonths(_cobStartMonth(lastPago), 1).getTime() - 86400000) : null;
    proximoCobro = cobertoHastaMes;
    if (mesesDebe <= 0) {
      if (proximoCobro && proximoCobro.getTime() <= finVentana.getTime()) { estatus = 'Por vencer'; montoDebe = cfg.tarifaMensual; }
      else estatus = 'Vigente';
    } else {
      estatus = 'Vencida';
      montoDebe = mesesDebe * cfg.tarifaMensual;
    }
  } else {
    // ANUAL (o default anual)
    var coberturaFin; // fecha hasta la que está cubierto (exclusivo → primer día NO cubierto)
    if (lastPago) {
      coberturaFin = _cobAddYears(lastPago, 1);
      if (coberturaFin.getTime() < billStart.getTime()) coberturaFin = billStart;
    } else {
      coberturaFin = billStart;  // nunca pagó → debe desde billStart
    }
    coberturaHasta = lastPago ? new Date(_cobAddYears(lastPago, 1).getTime() - 86400000) : null;
    proximoCobro = coberturaFin;
    if (today.getTime() < coberturaFin.getTime()) {
      if (coberturaFin.getTime() <= finVentana.getTime()) { estatus = 'Por vencer'; montoDebe = cfg.tarifaAnual; }
      else { estatus = 'Vigente'; montoDebe = 0; }
    } else {
      // años vencidos
      var aniosDebe = 0, cursor = new Date(coberturaFin.getTime());
      while (cursor.getTime() <= today.getTime() && aniosDebe < 20) { aniosDebe++; cursor = _cobAddYears(cursor, 1); }
      aniosDebe = Math.max(1, aniosDebe);
      montoDebe = aniosDebe * cfg.tarifaAnual;
      mesesDebe = aniosDebe * 12;
      estatus = 'Vencida';
    }
    if (!tienePlan && crio && today.getTime() >= billStart.getTime()) {
      // tiene crío, sin plan elegido y ya pasó la cortesía → falta suscripción
      estatus = (montoDebe > 0) ? 'Falta suscripción' : estatus;
    }
  }

  if (!tienePlan && estatus !== 'Cortesía' && estatus !== 'Por vencer' && montoDebe > 0) estatus = 'Falta suscripción';

  return {
    plan: plan, tarifa: tarifa, esExterno: esExterno,
    crioInicio: _cobStr(crio), billStart: _cobStr(billStart),
    ultimoPago: _cobStr(lastPago), ultProducto: ultProducto,
    coberturaHasta: _cobStr(coberturaHasta), proximoCobro: _cobStr(proximoCobro),
    estatus: estatus, montoDebe: montoDebe, mesesDebe: mesesDebe, tienePlan: tienePlan,
    oov: pac.oov, emb: pac.emb
  };
}

function _cobBuildSuscripciones() {
  var cfg = _cobCfg();
  var today = _cobToday();
  var crioData = _cobReadCrio();
  if (crioData.error) return { error: crioData.error, pacientes: [] };
  var pagos = _cobStoragePagosByPac();
  var abonos = _cobReadAbonos();
  var ledger = _susReadLedger();
  var lista = [];
  for (var i = 0; i < crioData.pacientes.length; i++) {
    var pac = crioData.pacientes[i];
    var calc = _cobCalcSuscripcion(pac, pagos[pac.key] || [], abonos.susByPac[pac.key] || [], cfg, today);
    // Materializado (generado) pendiente de este paciente — informativo; el adeudo
    // real sigue siendo calc.montoDebe (una sola fuente, sin doble conteo).
    var lrows = ledger.byPac[pac.key] || [];
    var genPend = 0, genCount = 0;
    for (var lr = 0; lr < lrows.length; lr++) {
      var est = _cobLower(lrows[lr].estatus);
      if (est !== 'pagado' && est !== 'cancelado') { genPend += (lrows[lr].monto - lrows[lr].abonado); genCount++; }
    }
    lista.push({
      nombre: _cobMasked() ? ('Paciente ' + (i + 1)) : pac.nombre,
      plan: calc.plan, tipoPaciente: calc.esExterno ? 'Externo' : 'Hestia',
      crioInicio: calc.crioInicio, billStart: calc.billStart,
      ultimoPago: calc.ultimoPago, coberturaHasta: calc.coberturaHasta, proximoCobro: calc.proximoCobro,
      estatus: calc.estatus, montoDebe: calc.montoDebe, mesesDebe: calc.mesesDebe,
      generado: genPend, generadoCount: genCount, porGenerar: Math.max(0, calc.montoDebe - genPend),
      oov: calc.oov, emb: calc.emb, incEmb: pac.incEmb || 0, tienePlan: calc.tienePlan,
      autopay: !!calc.autopay, mensual: !!pac.mensual
    });
  }
  // orden: vencidas primero (por monto), luego por vencer, luego el resto
  var rank = { 'Vencida': 0, 'Falta suscripción': 1, 'Por vencer': 2, 'Vigente': 3, 'Cortesía': 4, 'Pago automático': 5 };
  lista.sort(function (a, b) {
    var ra = rank[a.estatus] === undefined ? 5 : rank[a.estatus];
    var rb = rank[b.estatus] === undefined ? 5 : rank[b.estatus];
    if (ra !== rb) return ra - rb;
    return b.montoDebe - a.montoDebe;
  });
  return { pacientes: lista };
}

function readSuscripcionesCrio(body) {
  try {
    var built = _cobBuildSuscripciones();
    if (built.error) return { ok: false, error: built.error, version: COBRANZA_VER, pacientes: [] };
    var lista = built.pacientes;
    var tot = { total: 0, porEstatus: {}, montoPorEstatus: {}, totalDebe: 0, numConCrio: lista.length };
    for (var i = 0; i < lista.length; i++) {
      var p = lista[i];
      tot.porEstatus[p.estatus] = (tot.porEstatus[p.estatus] || 0) + 1;
      tot.montoPorEstatus[p.estatus] = (tot.montoPorEstatus[p.estatus] || 0) + p.montoDebe;
      tot.totalDebe += p.montoDebe;
    }
    return {
      ok: true, version: COBRANZA_VER, view: 'suscripciones-crio',
      pacientes: lista, totalDebe: tot.totalDebe, numConCrio: lista.length,
      porEstatus: tot.porEstatus, montoPorEstatus: tot.montoPorEstatus, masked: _cobMasked()
    };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER, pacientes: [] };
  }
}

/* ═════════════════ Estado de cuenta de cobranza (por paciente) ═════════════════ */
// Detalle COMPLETO de lo que se le vendió al paciente (todas sus líneas de BD_Ingresos).
function _cobVentasPaciente(keyBuscar) {
  var out = [];
  var sh = SpreadsheetApp.openById(INGRESOS_SS_ID).getSheetByName(BD_INGRESOS_TAB);
  if (!sh) return out;
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return out;
  var H = data[0].map(function (x) { return _cobLower(x); });
  var hc = function () { for (var a = 0; a < arguments.length; a++) { var k = H.indexOf(arguments[a]); if (k > -1) return k; } return -1; };
  var iOp = hc('op'), iFecha = hc('fecha'), iPac = hc('paciente'), iCat = hc('categoria', 'categoría'),
      iProd = hc('producto'), iCant = hc('cantidad'), iTotal = hc('totalpagar', 'total a pagar', 'total'),
      iPag = hc('pagado'), iFP = hc('formapago', 'forma de pago', 'forma pago');
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (_cobKeyNom(row[iPac]) !== keyBuscar) continue;
    var total = _cobNum(row[iTotal]);
    var pag = iPag > -1 ? _cobNum(row[iPag]) : 0;
    var saldoL = (pag > 0.01 && pag < total - 0.01) ? (total - pag) : 0; // saldo real solo si pago parcial
    out.push({
      fecha: _cobStr(_cobD(row[iFecha])), op: String(row[iOp] || '').trim(),
      producto: String(row[iProd] || ''), categoria: String(row[iCat] || ''),
      cantidad: iCant > -1 ? _cobNum(row[iCant]) : 1, total: total, pagado: pag, saldo: saldoL,
      formaPago: iFP > -1 ? String(row[iFP] || '') : ''
    });
  }
  out.sort(function (a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });
  return out;
}
function readEstadoCobranza(pacienteNombre) {
  try {
    if (!pacienteNombre) return { ok: false, error: 'Nombre de paciente requerido' };
    var keyBuscar = _cobKeyNom(pacienteNombre);

    // Motor A: saldos de este paciente
    var saldos = _cobBuildSaldos().ops.filter(function (o) {
      return _cobKeyNom(o.paciente) === keyBuscar || o.paciente === pacienteNombre;
    });
    // Si está enmascarado, _cobBuildSaldos ya devolvió OP-#### y no se puede filtrar por nombre;
    // en ese caso no exponemos detalle por nombre.
    var totalSaldo = 0; saldos.forEach(function (o) { totalSaldo += o.saldo; });

    // Abonos de este paciente
    var abonos = _cobReadAbonos();
    var misAbonos = abonos.rows.filter(function (a) { return _cobKeyNom(a.paciente) === keyBuscar; });

    // Motor B: suscripción de este paciente
    var susList = _cobBuildSuscripciones().pacientes || [];
    var miSus = null;
    for (var i = 0; i < susList.length; i++) {
      if (_cobKeyNom(susList[i].nombre) === keyBuscar) { miSus = susList[i]; break; }
    }
    // Suscripciones materializadas (generadas) pendientes de este paciente
    var ledger = _susReadLedger();
    var lrows = (ledger.byPac[keyBuscar] || []).filter(function (r) {
      var e = _cobLower(r.estatus); return e !== 'pagado' && e !== 'cancelado';
    });
    var susCargos = lrows.map(function (r) {
      return { fecha: r.fecha, concepto: r.concepto, periodo: r.periodo, monto: r.monto - r.abonado, susId: r.susId };
    });
    var susCargosTotal = 0; susCargos.forEach(function (r) { susCargosTotal += r.monto; });

    // Detalle completo de lo vendido (todas las líneas de ingresos del paciente)
    var ventas = _cobMasked() ? [] : _cobVentasPaciente(keyBuscar);
    var totalVendido = 0, totalPagadoVentas = 0;
    ventas.forEach(function (v) { totalVendido += v.total; totalPagadoVentas += v.pagado; });

    return {
      ok: true, version: COBRANZA_VER,
      paciente: _cobMasked() ? '—' : String(pacienteNombre).trim(),
      masked: _cobMasked(),
      saldos: saldos, totalSaldo: totalSaldo,
      abonos: misAbonos, suscripcion: miSus,
      suscripcionCargos: susCargos, suscripcionCargosTotal: susCargosTotal,
      ventas: ventas, totalVendido: totalVendido, totalPagadoVentas: totalPagadoVentas
    };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}

/* ═════════════════ Escritura: abonos y saldos iniciales ═════════════════ */
function registrarAbono(body) {
  try {
    var monto = _cobNum(body.monto);
    var tipo = _cobLower(body.tipo) || 'saldo';
    if (monto <= 0) return { ok: false, error: 'Monto inválido' };
    var sh = _cobEnsureAbonos();
    var fecha = body.fecha ? _cobStr(_cobD(body.fecha)) : _cobStr(_cobToday());
    sh.appendRow([
      fecha, String(body.op || '').trim(), String(body.paciente || '').trim(), monto,
      String(body.formaPago || ''), String(body.banco || ''), tipo,
      String(body.nota || ''), String(body.usuario || ''), new Date()
    ]);
    // Cobro de suscripción: marca los periodos generados (ledger) como pagados FIFO.
    if (tipo === 'suscripcion' && body.paciente) { _susAplicarPagoFIFO(body.paciente, monto); }
    return { ok: true, version: COBRANZA_VER, fecha: fecha, monto: monto };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}

function cargarSaldoInicial(body) {
  try {
    var items = _cobParseItems(body.items);
    var monto = items.length ? _cobItemsMonto(items) : _cobNum(body.monto);
    if (monto <= 0) return { ok: false, error: 'Captura al menos una partida (producto) o un monto.' };
    var sh = _cobEnsureCargos();
    var iItems = _cobCargosItemsCol(sh);
    var fecha = body.fecha ? _cobStr(_cobD(body.fecha)) : _cobStr(_cobToday());
    var concepto = String(body.concepto || '').trim();
    if (!concepto) concepto = items.length ? (items.map(function (it) { return it.producto; }).filter(Boolean).slice(0, 3).join(', ') || 'Cuenta por cobrar') : 'Saldo inicial';
    sh.appendRow([
      fecha, String(body.op || '').trim(), String(body.paciente || '').trim(),
      String(body.categoria || ''), concepto,
      monto, String(body.estatus || 'Pendiente'), String(body.nota || ''),
      String(body.usuario || ''), new Date()
    ]);
    if (items.length) sh.getRange(sh.getLastRow(), iItems).setValue(JSON.stringify(items));
    return { ok: true, version: COBRANZA_VER, fecha: fecha, monto: monto };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}

/* ═════════════ GENERACIÓN DE SUSCRIPCIONES (materializar por periodo) ═════════════
 * Espejo de Gastos Fijos: cada periodo vencido se materializa como un renglón en la
 * hoja Suscripciones_Crio (llave anti-duplicado paciente|plan|periodo + LockService).
 * El backlog histórico se consolida en UN solo renglón 'SALDO-INICIAL' por paciente;
 * de ahí en adelante, uno por periodo. Idempotente: correrla N veces no duplica.
 */
function _susEnsureSheet() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_SUS);
  if (!sh) {
    sh = ss.insertSheet(COBRANZA_SUS);
    sh.appendRow(['SuscripcionID','Fecha','Paciente','Plan','PeriodoInicio','PeriodoFin','Periodo','Monto','Estatus','Abonado','Concepto','Usuario','Timestamp']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function _susReadLedger() {
  var out = { byPac: {}, keys: {}, byId: {} };
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_SUS);
  if (!sh) return out;
  var raw = sh.getDataRange().getValues();
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var susId = String(r[0] || '').trim(); if (!susId) continue;
    var pacKey = _cobKeyNom(r[2]);
    var periodo = String(r[6] || '').trim();
    var rec = {
      rowNum: i + 1, susId: susId, fecha: _cobStr(_cobD(r[1])), paciente: String(r[2] || ''),
      plan: String(r[3] || ''), inicio: _cobStr(_cobD(r[4])), fin: _cobStr(_cobD(r[5])),
      periodo: periodo, monto: _cobNum(r[7]), estatus: String(r[8] || 'Pendiente'),
      abonado: _cobNum(r[9]), concepto: String(r[10] || '')
    };
    if (!out.byPac[pacKey]) out.byPac[pacKey] = [];
    out.byPac[pacKey].push(rec);
    if (!out.keys[pacKey]) out.keys[pacKey] = {};
    out.keys[pacKey][periodo] = true;
    out.byId[susId] = rec;
  }
  return out;
}
// Periodos vencidos (no cubiertos por un pago de almacenamiento en Ingresos) desde
// billStart hasta hoy. Cada uno: {key, inicio, fin, vencimiento, monto}.
function _susPeriodosDue(pac, calc, cfg, today) {
  var out = [];
  var billStart = _cobD(calc.billStart); if (!billStart) return out;
  var lastPago = _cobD(calc.ultimoPago);
  if (calc.plan === 'mensual') {
    var first = _cobStartMonth(billStart);
    if (lastPago) { var afterPay = _cobStartMonth(_cobAddMonths(lastPago, 1)); if (afterPay.getTime() > first.getTime()) first = afterPay; }
    var cur = first, g = 0;
    while (cur.getTime() <= today.getTime() && g < cfg.maxMesesAtraso + 2) {
      var fin = new Date(_cobAddMonths(cur, 1).getTime() - 86400000);
      out.push({ key: _cobStr(cur).substring(0, 7), inicio: _cobStr(cur), fin: _cobStr(fin), vencimiento: _cobStr(cur), monto: cfg.tarifaMensual });
      cur = _cobAddMonths(cur, 1); g++;
    }
  } else {
    var firstY = billStart;
    if (lastPago) { var afterPayY = _cobAddYears(lastPago, 1); if (afterPayY.getTime() > firstY.getTime()) firstY = afterPayY; }
    var curY = firstY, g2 = 0;
    while (curY.getTime() <= today.getTime() && g2 < 40) {
      var finY = new Date(_cobAddYears(curY, 1).getTime() - 86400000);
      out.push({ key: _cobStr(curY).substring(0, 7), inicio: _cobStr(curY), fin: _cobStr(finY), vencimiento: _cobStr(curY), monto: cfg.tarifaAnual });
      curY = _cobAddYears(curY, 1); g2++;
    }
  }
  return out;
}
function _susRowFromPeriod(pac, calc, p) {
  var lbl;
  if (calc.plan === 'mensual') lbl = 'Suscripción crío mensual ' + p.key;
  else { var y1 = p.inicio.substring(0, 4), y2 = p.fin.substring(0, 4); lbl = 'Suscripción crío anual ' + y1 + (y2 !== y1 ? '–' + y2 : ''); }
  return { pacKey: pac.key, paciente: pac.nombre, plan: calc.plan, periodo: p.key,
    inicio: p.inicio, fin: p.fin, monto: p.monto, vencimiento: p.vencimiento, concepto: lbl };
}
function generarSuscripciones(body) {
  try {
    body = body || {};
    var cfg = _cobCfg(), today = _cobToday();
    var preview = !!body.preview;
    var soloPac = body.soloPaciente ? _cobKeyNom(body.soloPaciente) : null;
    var consolidar = !(body.consolidar === false);
    var crioData = _cobReadCrio();
    if (crioData.error) return { ok: false, error: crioData.error, version: COBRANZA_VER };
    var pagos = _cobStoragePagosByPac();
    var abonos = _cobReadAbonos();
    var ledger = _susReadLedger();
    var plan = [];
    for (var i = 0; i < crioData.pacientes.length; i++) {
      var pac = crioData.pacientes[i];
      if (soloPac && pac.key !== soloPac) continue;
      var calc = _cobCalcSuscripcion(pac, pagos[pac.key] || [], abonos.susByPac[pac.key] || [], cfg, today);
      if (!calc.crioInicio || calc.estatus === 'Cortesía') continue;
      var due = _susPeriodosDue(pac, calc, cfg, today);
      if (!due.length) continue;
      var existing = ledger.keys[pac.key] || {};
      var news = [];
      for (var d = 0; d < due.length; d++) { if (!existing[due[d].key]) news.push(due[d]); }
      if (!news.length) continue;
      var curKey = due[due.length - 1].key;
      var backlog = news.filter(function (p) { return p.key !== curKey; });
      var current = news.filter(function (p) { return p.key === curKey; });
      var hasConsol = !!existing['SALDO-INICIAL'];
      if (backlog.length) {
        if (consolidar && !hasConsol) {
          var sum = 0; backlog.forEach(function (p) { sum += p.monto; });
          plan.push({ pacKey: pac.key, paciente: pac.nombre, plan: calc.plan, periodo: 'SALDO-INICIAL',
            inicio: backlog[0].inicio, fin: backlog[backlog.length - 1].fin, monto: sum,
            vencimiento: backlog[0].vencimiento, concepto: 'Saldo inicial suscripción — ' + backlog.length + ' periodo(s) al ' + _cobStr(today) });
        } else {
          backlog.forEach(function (p) { plan.push(_susRowFromPeriod(pac, calc, p)); });
        }
      }
      current.forEach(function (p) { plan.push(_susRowFromPeriod(pac, calc, p)); });
    }
    if (preview) {
      var tot = 0; plan.forEach(function (r) { tot += r.monto; });
      return { ok: true, version: COBRANZA_VER, preview: true, count: plan.length, monto: tot,
        consolidados: plan.filter(function (r) { return r.periodo === 'SALDO-INICIAL'; }).length,
        detalle: plan.slice(0, 300).map(function (r) { return { paciente: _cobMasked() ? '—' : r.paciente, plan: r.plan, periodo: r.periodo, monto: r.monto, concepto: r.concepto }; }) };
    }
    // COMMIT (idempotente, con lock) — requiere permiso (afecta datos)
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_egresos'))
      return { ok: false, error: 'Sin permiso para generar suscripciones (requiere editar cuentas por pagar).', version: COBRANZA_VER };
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado, reintenta.', version: COBRANZA_VER };
    var creadas = 0, dups = 0, montoG = 0;
    try {
      var sh = _susEnsureSheet();
      var led2 = _susReadLedger();
      for (var j = 0; j < plan.length; j++) {
        var r = plan[j];
        var ex = led2.keys[r.pacKey] || {};
        if (ex[r.periodo]) { dups++; continue; }
        var susId = r.pacKey + '|' + r.plan + '|' + r.periodo;
        sh.appendRow([susId, r.vencimiento, r.paciente, r.plan, r.inicio, r.fin, r.periodo, r.monto, 'Pendiente', 0, r.concepto, (body.usuario || 'sistema'), new Date()]);
        if (!led2.keys[r.pacKey]) led2.keys[r.pacKey] = {};
        led2.keys[r.pacKey][r.periodo] = true;
        creadas++; montoG += r.monto;
      }
    } finally { lock.releaseLock(); }
    return { ok: true, version: COBRANZA_VER, creadas: creadas, duplicadas: dups, monto: montoG };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}
// Aplica un cobro de suscripción a los periodos generados pendientes (FIFO): marca
// filas del ledger como Pagado/parcial hasta agotar el monto. Mantiene el ledger en
// sincronía con el cobro (la cobertura avanza vía el abono en Abonos_Cobrar).
function _susAplicarPagoFIFO(paciente, monto) {
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(COBRANZA_SUS); if (!sh) return;
    var key = _cobKeyNom(paciente);
    var raw = sh.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < raw.length; i++) {
      if (_cobKeyNom(raw[i][2]) !== key) continue;
      var est = _cobLower(raw[i][8]);
      if (est === 'pagado' || est === 'cancelado') continue;
      rows.push({ rowNum: i + 1, fecha: _cobD(raw[i][1]), monto: _cobNum(raw[i][7]), abonado: _cobNum(raw[i][9]) });
    }
    rows.sort(function (a, b) { return (a.fecha ? a.fecha.getTime() : 0) - (b.fecha ? b.fecha.getTime() : 0); });
    var rem = monto;
    for (var k = 0; k < rows.length && rem > 0.01; k++) {
      var row = rows[k];
      var pend = row.monto - row.abonado;
      var apply = Math.min(pend, rem);
      var nuevoAb = row.abonado + apply;
      sh.getRange(row.rowNum, 10).setValue(nuevoAb);           // Abonado
      if (nuevoAb >= row.monto - 0.01) sh.getRange(row.rowNum, 9).setValue('Pagado');
      rem -= apply;
    }
  } catch (e) {}
}
// Handler para el scheduler (genera lo del mes automáticamente, sin consolidar backlog).
function generarSuscripcionesMensual() {
  return generarSuscripciones({ usuario: 'sistema (scheduler)', consolidar: false });
}
