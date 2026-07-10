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
var COBRANZA_VER      = 'cobranza-2026.07.09a';

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
  return { ok: true, version: COBRANZA_VER, config: cfg };
}

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
    sh.appendRow(['Fecha', 'OP', 'Paciente', 'Categoria', 'Concepto', 'MontoCargo', 'Estatus', 'Nota', 'Usuario', 'Timestamp']);
    sh.setFrozenRows(1);
  }
  return sh;
}
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
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var op = String(r[1] || '').trim();
    var monto = _cobNum(r[5]);
    if (monto <= 0) continue;
    var rec = {
      fecha: _cobStr(_cobD(r[0])), op: op, paciente: String(r[2] || ''),
      categoria: String(r[3] || ''), concepto: String(r[4] || ''),
      monto: monto, estatus: String(r[6] || ''), nota: String(r[7] || '')
    };
    out.rows.push(rec);
    var key = op || ('_cargo_' + i);
    out.byOp[key] = rec;
  }
  return out;
}

/* ═════════════════ MOTOR A — Saldos por cobrar ═════════════════ */
function _cobBuildSaldos() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(BD_INGRESOS_TAB);
  if (!sh) return { ops: [], error: 'No se encontró ' + BD_INGRESOS_TAB };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ops: [] };

  var H = data[0].map(function (x) { return _cobLower(x); });
  function hc() { for (var a = 0; a < arguments.length; a++) { var k = H.indexOf(arguments[a]); if (k > -1) return k; } return -1; }
  var iOp    = hc('op'); if (iOp < 0) iOp = 0;
  var iFecha = hc('fecha');
  var iPac   = hc('paciente');
  var iCat   = hc('categoria', 'categoría');
  var iProd  = hc('producto');
  var iTotal = hc('totalpagar', 'total a pagar', 'total');
  var iPag   = hc('pagado');

  var abonos = _cobReadAbonos();
  var cargos = _cobReadCargos();

  var ops = {};  // op → agregado
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var op = String(row[iOp] || '').trim();
    if (!op) continue;
    var total = _cobNum(row[iTotal]);
    var pagCell = iPag > -1 ? row[iPag] : '';
    var pagBlankRow = _cobIsBlank(pagCell);
    var pag = _cobNum(pagCell);
    if (!ops[op]) {
      ops[op] = {
        op: op, paciente: String(row[iPac] || '').trim(),
        categoria: String(row[iCat] || '').trim(), producto: String(row[iProd] || '').trim(),
        fecha: _cobStr(_cobD(row[iFecha])), total: 0, pagadoCol: 0, algunPagado: false, conceptos: []
      };
    }
    var o = ops[op];
    o.total += total;
    if (!pagBlankRow) { o.pagadoCol += pag; o.algunPagado = true; }
    if (o.conceptos.length < 6 && o.producto) o.conceptos.push(o.producto);
    if (!o.fecha) o.fecha = _cobStr(_cobD(row[iFecha]));
  }

  var today = _cobToday();
  var out = [];
  var seen = {};
  for (var k in ops) {
    var it = ops[k];
    // base a deber (antes de abonos):
    var base;
    if (cargos.byOp[it.op]) {
      base = cargos.byOp[it.op].monto;                 // registro manual explícito
    } else if (it.algunPagado && it.pagadoCol < it.total - 0.01) {
      base = it.total - it.pagadoCol;                  // pago parcial capturado en Ingresos
    } else {
      base = 0;                                        // histórico/sin info = pagado
    }
    var abo = abonos.byOp[it.op] || 0;
    var saldo = base - abo;
    if (saldo <= 0.01) continue;
    seen[it.op] = true;
    var fd = _cobD(it.fecha) || today;
    var dias = _cobDaysDiff(fd, today); if (dias < 0) dias = 0;
    out.push({
      op: it.op,
      paciente: _cobMasked() ? _privPaciente(it.op) : it.paciente,
      segmento: _cobSegmento(it.categoria, it.producto),
      categoria: it.categoria,
      concepto: it.conceptos.join(', ') || it.producto,
      fecha: it.fecha,
      total: it.total, abonado: (it.algunPagado ? it.pagadoCol : 0) + abo, saldo: saldo,
      dias: dias, bucket: _cobBucket(dias)
    });
  }
  // cargos manuales cuyo OP no está en Ingresos (ej. saldos iniciales sin OP)
  for (var ci = 0; ci < cargos.rows.length; ci++) {
    var cg = cargos.rows[ci];
    if (cg.op && seen[cg.op]) continue;               // ya contado arriba
    if (cg.op && ops[cg.op]) continue;
    var abo2 = cg.op ? (abonos.byOp[cg.op] || 0) : 0;
    var saldo2 = cg.monto - abo2;
    if (saldo2 <= 0.01) continue;
    var fd2 = _cobD(cg.fecha) || today;
    var dias2 = _cobDaysDiff(fd2, today); if (dias2 < 0) dias2 = 0;
    out.push({
      op: cg.op || '—',
      paciente: _cobMasked() ? (cg.op ? _privPaciente(cg.op) : 'Paciente') : cg.paciente,
      segmento: _cobSegmento(cg.categoria, cg.concepto),
      categoria: cg.categoria, concepto: cg.concepto || cg.estatus || 'Saldo inicial',
      fecha: cg.fecha, total: cg.monto, abonado: abo2, saldo: saldo2,
      dias: dias2, bucket: _cobBucket(dias2)
    });
  }
  out.sort(function (a, b) { return b.saldo - a.saldo; });
  return { ops: out };
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
    var extRaw = iExt > -1 ? row[iExt] : '';
    var esExterno = _cobEsExterno(extRaw);
    if (!map[key]) map[key] = { nombre: nom, key: key, crioInicio: f, oov: 0, emb: 0, externo: false, _extSet: false };
    var m = map[key];
    if (f && (!m.crioInicio || f.getTime() < m.crioInicio.getTime())) m.crioInicio = f;
    m.oov += oov; m.emb += emb;
    if (iExt > -1 && !_cobIsBlank(extRaw)) { m.externo = m.externo || esExterno; m._extSet = true; }
  }
  for (var k in map) {
    var mm = map[k];
    if ((mm.oov + mm.emb) <= 0 && !mm.crioInicio) continue;  // sin inventario ni fecha → ignorar
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
  // plan: del último pago de almacenamiento que tenga plan definido, si no default
  var plan = cfg.planDefault, meses = (cfg.planDefault === 'mensual' ? 1 : 12), tarifa = (cfg.planDefault === 'mensual' ? cfg.tarifaMensual : cfg.tarifaAnual);
  var lastPago = null, ultProducto = '';
  var todos = (pagos || []).slice();
  // incluir abonos de suscripción como "pagos"
  (susAbonos || []).forEach(function (a) { var f = _cobD(a.fecha); if (f) todos.push({ fecha: f, producto: 'Abono suscripción', monto: a.monto, plan: null, meses: null }); });
  todos.sort(function (a, b) { return a.fecha.getTime() - b.fecha.getTime(); });
  for (var i = 0; i < todos.length; i++) {
    var p = todos[i];
    if (p.plan) { plan = p.plan; meses = p.meses; ultProducto = p.producto; if (plan === 'mensual') tarifa = cfg.tarifaMensual; else if (plan === 'anual') tarifa = cfg.tarifaAnual; else tarifa = p.monto || tarifa; }
    if (!lastPago || p.fecha.getTime() > lastPago.getTime()) lastPago = p.fecha;
  }
  var tienePlan = todos.length > 0;

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

  // ¿En cortesía (año gratis) y sin haber pagado aún?
  if (crio && !esExterno && cfg.anioGratisSoloHestia && today.getTime() < billStart.getTime() && !lastPago) {
    estatus = 'Cortesía';
    proximoCobro = billStart;
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
      estatus = (proximoCobro && _cobDaysDiff(today, proximoCobro) <= cfg.diasPorVencer) ? 'Por vencer' : 'Vigente';
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
      estatus = (_cobDaysDiff(today, coberturaFin) <= cfg.diasPorVencer) ? 'Por vencer' : 'Vigente';
      montoDebe = 0;
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

  if (!tienePlan && estatus !== 'Cortesía' && montoDebe > 0) estatus = 'Falta suscripción';

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
  var lista = [];
  for (var i = 0; i < crioData.pacientes.length; i++) {
    var pac = crioData.pacientes[i];
    var calc = _cobCalcSuscripcion(pac, pagos[pac.key] || [], abonos.susByPac[pac.key] || [], cfg, today);
    lista.push({
      nombre: _cobMasked() ? ('Paciente ' + (i + 1)) : pac.nombre,
      plan: calc.plan, tipoPaciente: calc.esExterno ? 'Externo' : 'Hestia',
      crioInicio: calc.crioInicio, billStart: calc.billStart,
      ultimoPago: calc.ultimoPago, coberturaHasta: calc.coberturaHasta, proximoCobro: calc.proximoCobro,
      estatus: calc.estatus, montoDebe: calc.montoDebe, mesesDebe: calc.mesesDebe,
      oov: calc.oov, emb: calc.emb, tienePlan: calc.tienePlan
    });
  }
  // orden: vencidas primero (por monto), luego por vencer, luego el resto
  var rank = { 'Vencida': 0, 'Falta suscripción': 1, 'Por vencer': 2, 'Vigente': 3, 'Cortesía': 4 };
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

    return {
      ok: true, version: COBRANZA_VER,
      paciente: _cobMasked() ? '—' : String(pacienteNombre).trim(),
      masked: _cobMasked(),
      saldos: saldos, totalSaldo: totalSaldo,
      abonos: misAbonos, suscripcion: miSus
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
    return { ok: true, version: COBRANZA_VER, fecha: fecha, monto: monto };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}

function cargarSaldoInicial(body) {
  try {
    var monto = _cobNum(body.monto);
    if (monto <= 0) return { ok: false, error: 'Monto inválido' };
    var sh = _cobEnsureCargos();
    var fecha = body.fecha ? _cobStr(_cobD(body.fecha)) : _cobStr(_cobToday());
    sh.appendRow([
      fecha, String(body.op || '').trim(), String(body.paciente || '').trim(),
      String(body.categoria || ''), String(body.concepto || 'Saldo inicial'),
      monto, String(body.estatus || 'Saldo inicial'), String(body.nota || ''),
      String(body.usuario || ''), new Date()
    ]);
    return { ok: true, version: COBRANZA_VER, fecha: fecha, monto: monto };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}
