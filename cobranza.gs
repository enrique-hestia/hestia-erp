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
var COBRANZA_DESCUENTOS = 'Descuentos_Agencia'; // escala de descuento por volumen, por agencia
var COBRANZA_DEPOSITO_KEY = 'COBRANZA_DEPOSITO';
var COBRANZA_VER      = 'cobranza-2026.07.15a';
// Cuenta de depósito por defecto (Hestia recibe aquí). Editable en Script Property COBRANZA_DEPOSITO.
var COBRANZA_DEPOSITO_DEF = { banco:'Santander', beneficiario:'Hestia Clinic', cuenta:'65-51043096-7', clabe:'014180655104309670' };

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
  if (typeof _mpEnsureSheet === 'function') _mpEnsureSheet();
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
    // La cantidad tiene DOS usos y no se pueden mezclar:
    //  · cRaw = lo que de verdad se capturó. null = no se capturó. Es lo que se
    //    IMPRIME, y por eso no puede llevar un 1 inventado: un documento que le
    //    llega a una agencia no puede afirmar "1" donde nadie escribió nada.
    //  · c    = el 1 solo para la ARITMÉTICA (p × c). Si no hay cantidad, el
    //    precio ES el total de la partida; asumir 1 aquí es correcto y necesario
    //    (ponerle null daría p × null = 0 y borraría el importe).
    var cRaw = (it.cantidad == null || it.cantidad === '') ? null : _cobNum(it.cantidad);
    var c = cRaw || 1;
    var p = _cobNum(it.precio);
    var t = (it.total != null && it.total !== '') ? _cobNum(it.total) : (p * c);
    var pac = String(it.pac || it.paciente || '').trim(); // nombre del paciente real (agencias/externos)
    var fecha = String(it.fecha || '').trim();            // fecha del tratamiento de esa partida
    if (String(it.producto || '').trim() || t > 0) out.push({ producto: String(it.producto || '').trim(), cantidad: cRaw, precio: p, total: t, saldo: 0, pac: pac, fecha: fecha });
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
      // 'abono-op'    = abono del botón "Cobrar / Abonar" de una OP (finance.abonarIngreso).
      // 'abono-cargo' = abono a un cargo explícito de Cuentas_Cobrar (cobranza.abonarCargo).
      // Ambos flujos YA dejaron el saldo canónico escrito en el MontoCargo del renglón de
      // Cuentas_Cobrar ('abono-op' lo recomputa como Facturado − Pagado; 'abono-cargo' lo
      // decrementa directo). Por eso NO se restan otra vez por OP aquí: sería doble conteo
      // del mismo cobro. Quedan en 'rows' (historial/traza y "Mis abonos").
      if (op && tipo !== 'abono-op' && tipo !== 'abono-cargo') out.byOp[op] = (out.byOp[op] || 0) + monto;
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
// Registra/actualiza automáticamente una cuenta por cobrar cuando un ingreso se
// captura o EDITA con pago parcial. La llaman saveIngreso y updateIngreso
// (finance.gs). Nunca debe tumbar la venta.
// IDEMPOTENTE (upsert por OP): la fila auto-generada se marca con Nota='auto-ingreso';
// editar el ingreso N veces reescribe ese MISMO renglón (no apila duplicados). Si el
// saldo baja a 0 (ya se pagó completo), el renglón se cierra (Estatus 'Pagado',
// Monto 0) sin borrarlo → queda como histórico y sale de Cuentas por Cobrar.
// 'monto' aquí es el SALDO (TotalPagar − Pagado) de la OP, no el total.
//
// DEVUELVE lo que de verdad hizo (antes no devolvía nada y el caller no podía saberlo):
//   { ok:true, accion:'actualizado'|'cerrado'|'creado'|'sin-cambio', row:<fila|0> }
//   { ok:false, error:'…' }
// 'sin-cambio' = no había renglón auto de esta OP y el saldo era 0 → NO HAY NADA que
// crear ni que cerrar. Es el caso que hacía que _canCerrarCxCobrar reportara éxito sin
// cerrar nada. Sigue sin tumbar la venta: los errores se atrapan y se reportan, no se
// avientan (todos los callers la llaman dentro de try/catch y descartan el resultado,
// así que devolver un objeto es retrocompatible).
function _cobRegistrarSaldoIngreso(op, paciente, categoria, monto, fecha) {
  try {
    op = String(op || '').trim();
    var saldo = _cobNum(monto);
    var sh = _cobEnsureCargos();
    var raw = sh.getDataRange().getValues();
    var hdr = (raw[0] || []).map(function (x) { return String(x).trim().toLowerCase(); });
    var iOp    = hdr.indexOf('op');        if (iOp < 0) iOp = 1;
    var iCat   = hdr.indexOf('categoria'); if (iCat < 0) iCat = 3;
    var iConc  = hdr.indexOf('concepto');  if (iConc < 0) iConc = 4;
    var iMonto = hdr.indexOf('montocargo');if (iMonto < 0) iMonto = 5;
    var iEst   = hdr.indexOf('estatus');   if (iEst < 0) iEst = 6;
    var iNota  = hdr.indexOf('nota');      if (iNota < 0) iNota = 7;
    var iPac   = hdr.indexOf('paciente');  if (iPac < 0) iPac = 2;
    // Localiza la fila auto-generada de ESTA OP (idempotencia).
    var foundRow = -1;
    if (op) {
      for (var r = 1; r < raw.length; r++) {
        if (String(raw[r][iOp] || '').trim() === op &&
            String(raw[r][iNota] || '').toLowerCase().indexOf('auto-ingreso') > -1) { foundRow = r + 1; break; }
      }
    }
    if (foundRow > 0) {
      if (saldo > 0.01) {
        sh.getRange(foundRow, iMonto + 1).setValue(saldo);
        sh.getRange(foundRow, iEst + 1).setValue('Pendiente');
        sh.getRange(foundRow, iConc + 1).setValue('Saldo pendiente de OP ' + op);
        if (paciente) sh.getRange(foundRow, iPac + 1).setValue(String(paciente).trim());
        if (categoria) sh.getRange(foundRow, iCat + 1).setValue(String(categoria));
        return { ok: true, accion: 'actualizado', row: foundRow };
      }
      // Sin saldo → cerrar (no borrar): sale de Cuentas por Cobrar, queda traza.
      // Se AUDITA el cierre: si un saldo real (>0) se cierra, tiene que poder
      // rastrearse quién/cuándo/desde qué monto (para detectar el error tipo
      // OP-01243: un editar que mandó pagado≥total y borró un adeudo vigente).
      var _prevMonto = _cobNum(raw[foundRow - 1][iMonto]);
      sh.getRange(foundRow, iMonto + 1).setValue(0);
      sh.getRange(foundRow, iEst + 1).setValue('Pagado');
      try {
        if (_prevMonto > 0.01 && typeof logAudit === 'function') {
          var _actor = (typeof _postEmail !== 'undefined' && _postEmail) ? _postEmail : 'sistema';
          logAudit(_actor, 'Cobranza', 'Cerrar saldo (auto)', op, 'MontoCargo', '$' + _prevMonto.toFixed(2) + ' pendiente', 'Pagado ($0)');
        }
      } catch (eLg) {}
      return { ok: true, accion: 'cerrado', row: foundRow };
    }
    // No existe todavía → solo se crea si de verdad hay saldo.
    if (saldo > 0.01) {
      sh.appendRow([
        fecha ? _cobStr(_cobD(fecha)) : _cobStr(_cobToday()),
        op, String(paciente || '').trim(), String(categoria || ''),
        'Saldo pendiente de OP ' + op, saldo, 'Pendiente', 'auto-ingreso', '', new Date()
      ]);
      return { ok: true, accion: 'creado', row: sh.getLastRow() };
    }
    // Ni renglón que tocar ni saldo que registrar: NO SE HIZO NADA. Se dice, no se
    // finge (un 'cerrado' aquí es lo que hacía mentir a la pantalla al cancelar).
    return { ok: true, accion: 'sin-cambio', row: 0 };
  } catch (e) { return { ok: false, error: e.message }; }
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
    // ── Renglón que MANDA la venta (auto-ingreso): su MontoCargo NO es editable aquí. Es
    // un DERIVADO (Facturado − Pagado de la OP) que _cobRegistrarSaldoIngreso reescribe en
    // el próximo abono o edición del ingreso → el cambio se evaporaría. Peor: entre tanto
    // la deuda diría una cosa y la venta otra. Se corrige en la VENTA, que es la fuente. ──
    var _au = _cobRowAuto(sh, rn);
    if (_au.auto)
      return { ok: false, version: COBRANZA_VER,
        error: 'Este saldo lo manda la venta ' + _au.op + ': aquí no se puede editar (el monto se recalcula solo como Facturado − Pagado y tu cambio se perdería). Corrige la venta en Ingresos → editar la operación ' + _au.op + ' y el saldo se ajusta solo.' };
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
    // ── Borrar el renglón de una VENTA a crédito dejaría el REVENUE huérfano: la venta
    // seguiría viva en Ingresos (sumando al P&L) sin nadie a quién cobrarle, y volver a
    // capturar el cargo duplicaría el ingreso. Deshacer una venta no es borrar su deuda:
    // se deshace EN LA VENTA. Esto NO cambia (el bloqueo es correcto) — lo que cambia es
    // A DÓNDE se manda al usuario, que hasta ahora era siempre «⊘ Cancelar».
    //
    // ── POR QUÉ SE RUTEA POR "¿YA SE COBRÓ ALGO?" ────────────────────────────────
    // Mandar todo a «⊘ Cancelar» dejaba un CALLEJÓN SIN SALIDA: cancelar_ingresos es un
    // permiso que POR DISEÑO ningún rol trae (cancelacion.gs), así que quien capturaba mal
    // una cuenta por cobrar no podía ni borrarla ni cancelarla, y el revenue falso se
    // quedaba en el P&L hasta que interviniera un admin. Y el alta ni siquiera pide
    // permiso (saveIngreso no tiene gate): cualquiera podía crear lo que nadie podía
    // deshacer.
    // El criterio correcto es si EL DINERO SE MOVIÓ:
    //   · Pagado = 0 → no se cobró un peso, no hay banco que reversar, no hay reembolso
    //     que hacer. Es un TYPEO, no una devolución. Se deshace BORRANDO la venta
    //     (borrar_ingresos), permiso que ya existe y que los roles sí pueden traer.
    //   · Pagado > 0 → entró dinero: deshacerla implica REVERSARLO y devolvérselo a la
    //     paciente. Eso es cancelar, y sigue exigiendo cancelar_ingresos.
    // No se abre ningún boquete: cancelar_ingresos sigue guardando TODOS los casos en que
    // hubo dinero, y borrar_ingresos ya permitía borrar ventas (con reverso de banco).
    // No se concede un solo poder nuevo; se deja de esconder la salida que ya existía.
    var _au = _cobRowAuto(sh, rn);
    if (_au.auto) {
      var _pg = _cobOPPagado(_au.op);
      if (!_pg.determinado)
        return { ok: false, version: COBRANZA_VER, op: _au.op, ruta: 'revisar',
          error: 'Este saldo viene de la venta ' + _au.op + ', pero esa venta no está en el libro de Ingresos de este año '
            + '(puede ser de un año anterior). No se borra desde aquí: dejaría la venta cobrando en Ingresos sin cuenta por cobrar, '
            + 'y no puedo verificar si ya tiene cobros. Pídele al administrador que la revise en el libro del año que corresponda.' };
      if (_pg.pagado > 0.01)
        return { ok: false, version: COBRANZA_VER, op: _au.op, ruta: 'cancelar',
          error: 'Este saldo viene de la venta ' + _au.op + ', que ya tiene $' + _pg.pagado.toFixed(2)
            + ' cobrados: borrarlo aquí dejaría la venta cobrando en Ingresos sin cuenta por cobrar. '
            + 'Como ya entró dinero, deshacerla es CANCELARLA (reversa el cobro en bancos y deja el registro para devolvérselo a la paciente): '
            + 'Ingresos → operación ' + _au.op + ' → «⊘ Cancelar». Si no te aparece ese botón, pídele al administrador '
            + '(o al director) que la cancele él, o que te otorgue el permiso "cancelar_ingresos" en Panel de Control → Roles.' };
      return { ok: false, version: COBRANZA_VER, op: _au.op, ruta: 'borrar',
        error: 'Este saldo viene de la venta a crédito ' + _au.op + ': borrarlo aquí dejaría la venta cobrando en Ingresos sin cuenta por cobrar. '
          + 'Como todavía NO se ha cobrado nada de esa venta, no hay dinero que devolver: si la capturaste por error, '
          + 'bórrala completa en Ingresos → operación ' + _au.op + ' → «🗑» y el adeudo se va con ella. '
          + 'Si no te aparece ese botón, pídele al administrador el permiso "borrar_ingresos" en Panel de Control → Roles.' };
    }
    sh.deleteRow(rn);
    return { ok: true, version: COBRANZA_VER, rowNum: rn };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_VER }; }
}
/* ¿Cuánto se ha COBRADO ya de la venta <op>? Decide la ruta para deshacerla:
 * pagado 0 → typeo, se borra (borrar_ingresos); pagado > 0 → entró dinero, se cancela
 * (cancelar_ingresos, que además lo reversa en bancos).
 * Misma definición de "Pagado" que abonarIngreso y el estado de cuenta: celda VACÍA ⇒
 * línea asumida PAGADA POR COMPLETO (histórico); celda CAPTURADA (incluido el 0 de una
 * venta a crédito) ⇒ literal. Con la lectura ingenua `num(celda)`, una venta histórica
 * en blanco se vería como "pagado 0" y mandaría a borrar algo ya cobrado.
 *
 * `determinado:false` = NO SE PUDO SABER, que NO es lo mismo que "pagado 0". La OP puede
 * no estar en el libro del año en curso (INGRESOS_SS_ID es de un año; hay un libro por
 * año) o la hoja puede no leerse. En ese caso NO se afirma nada y el caller manda a
 * revisión humana en vez de adivinar una ruta. */
function _cobOPPagado(op) {
  try {
    op = String(op || '').trim();
    if (!op) return { determinado: false, total: 0, pagado: 0 };
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(BD_INGRESOS_TAB);
    if (!sh) return { determinado: false, total: 0, pagado: 0 };
    var data = sh.getDataRange().getValues();
    var H = (data[0] || []).map(function (x) { return _cobLower(x); });
    var iOp = H.indexOf('op'); if (iOp < 0) iOp = 0;
    var iTot = H.indexOf('totalpagar'); if (iTot < 0) iTot = H.indexOf('total');
    var iPag = H.indexOf('pagado');
    if (iPag < 0) return { determinado: false, total: 0, pagado: 0 };
    var total = 0, pagado = 0, existe = false;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][iOp] || '').trim() !== op) continue;
      existe = true;
      var t = _cobNum(iTot > -1 ? data[r][iTot] : 0);
      var raw = data[r][iPag];
      total += t;
      pagado += (String(raw == null ? '' : raw).trim() === '') ? t : _cobNum(raw);
    }
    // Sin filas en ESTE libro: la venta puede vivir en el libro de otro año. No se
    // concluye "no se ha cobrado nada" a partir de no haberla encontrado.
    if (!existe) return { determinado: false, total: 0, pagado: 0 };
    return { determinado: true, total: total, pagado: pagado };
  } catch (e) { return { determinado: false, total: 0, pagado: 0 }; }
}
/* ¿El renglón <rn> de Cuentas_Cobrar lo manda una venta? (Nota 'auto-ingreso' + OP —
 * mismo criterio que _cobBuildSaldos para decidir auto:true, y que abonarCargo para
 * rechazar el cobro por la vía equivocada.) */
function _cobRowAuto(sh, rn) {
  try {
    var lc = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).trim().toLowerCase(); });
    var iOp = hdr.indexOf('op');     if (iOp < 0) iOp = 1;
    var iNota = hdr.indexOf('nota'); if (iNota < 0) iNota = 7;
    var row = sh.getRange(rn, 1, 1, lc).getValues()[0];
    var op = String(row[iOp] || '').trim();
    return { auto: (String(row[iNota] || '').toLowerCase().indexOf('auto-ingreso') > -1 && !!op), op: op };
  } catch (e) { return { auto: false, op: '' }; }
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
/* ── ALINEAR FOLIOS (arreglar el "drift" de folios) ───────────────────────────
 * NO renumera BD_Ingresos (es la fuente de verdad). Detecta cargos auto-ingreso de
 * Cuentas_Cobrar cuyo OP apunta a una venta que NO es de ese paciente (folio corrido,
 * p.ej. cargo "Kaitlyn $50,500" apuntando a OP-01246 que en BD es de Claire) y propone
 * RE-APUNTARLOS a la venta correcta cruzando por paciente + saldo (Total−Pagado) +
 * concepto/producto. Modo preview (default) NO escribe; apply=true aplica solo las
 * filas re-verificadas en la MISMA corrida. Gated por «reorganizar_folios». */
function alinearFoliosCobranza(body){
  try{
    body = body || {};
    if (typeof _tokenHasPermission==='function' && !_tokenHasPermission(body.token||'', 'reorganizar_folios'))
      return { ok:false, error:'Sin autorización para reorganizar folios (permiso «reorganizar_folios»).' };
    var apply = body.apply === true || body.apply === 'true';
    var actor = body.usuario || 'sistema';
    function num(v){ return _cobNum(v); }
    function nkey(s){ return _cobKeyNom(String(s||'')); }

    // Anti-TOCTOU: en modo APPLY tomamos el lock ANTES de leer, para que el cómputo
    // (BD + cargos) y la escritura sean consistentes; en preview no hace falta.
    var lock = null;
    if(apply){ lock = LockService.getScriptLock(); if(!lock.tryLock(25000)) return { ok:false, error:'Sistema ocupado, reintenta en un momento.' }; }
    try{

    // 1) BD_Ingresos → mapa por OP (sumando líneas) + índice por paciente
    var iss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var bd = iss.getSheetByName(BD_INGRESOS_TAB);
    if (!bd) return { ok:false, error:'No se encontró BD_Ingresos' };
    var dv = bd.getDataRange().getValues();
    var H = dv[0].map(function(x){ return String(x||'').trim().toLowerCase(); });
    function hc(){ for(var a=0;a<arguments.length;a++){ var k=H.indexOf(arguments[a]); if(k>-1) return k; } return -1; }
    var iOp=hc('op'); if(iOp<0) iOp=0;
    var iPac=hc('paciente'), iTot=hc('totalpagar','total a pagar','total'), iPag=hc('pagado'), iProd=hc('producto');
    var byOp={};
    for(var r=1;r<dv.length;r++){
      var op=String(dv[r][iOp]||'').trim(); if(!op) continue;
      if(!byOp[op]) byOp[op]={op:op, paciente:'', total:0, pagado:0, prods:[]};
      var t=num(dv[r][iTot]);
      var rawP=dv[r][iPag];
      var p=(String(rawP==null?'':rawP).trim()==='')? t : num(rawP);   // vacío = pagado (histórico)
      byOp[op].total+=t; byOp[op].pagado+=p;
      if(!byOp[op].paciente && iPac>-1) byOp[op].paciente=String(dv[r][iPac]||'').trim();
      if(iProd>-1 && dv[r][iProd]) byOp[op].prods.push(String(dv[r][iProd]));
    }
    var idxPac={};
    Object.keys(byOp).forEach(function(op){
      var o=byOp[op]; o.saldo=Math.round((o.total-o.pagado)*100)/100;
      var k=nkey(o.paciente); if(!idxPac[k]) idxPac[k]=[]; idxPac[k].push(o);
    });

    // 2) Cargos auto-ingreso de Cuentas_Cobrar
    var cc=_cobEnsureCargos();
    var cv=cc.getDataRange().getValues();
    var CH=cv[0].map(function(x){ return String(x||'').trim().toLowerCase(); });
    function cch(){ for(var a=0;a<arguments.length;a++){ var k=CH.indexOf(arguments[a]); if(k>-1) return k; } return -1; }
    var cOp=cch('op'); if(cOp<0)cOp=1;
    var cPac=cch('paciente'); if(cPac<0)cPac=2;
    var cConc=cch('concepto'); if(cConc<0)cConc=4;
    var cMonto=cch('montocargo'); if(cMonto<0)cMonto=5;
    var cEst=cch('estatus'); if(cEst<0)cEst=6;
    var cNota=cch('nota'); if(cNota<0)cNota=7;

    // folios BD que YA tienen un cargo alineado → no re-apuntar a uno ocupado.
    // Salta pagados/cancelados: un cargo cerrado NO debe reservar la venta de otro vivo.
    var ocupados={};
    for(var q=1;q<cv.length;q++){
      if(String(cv[q][cNota]||'').toLowerCase().indexOf('auto-ingreso')<0) continue;
      var eq=String(cv[q][cEst]||'').toLowerCase();
      if(eq==='pagado'||eq==='cancelado') continue;
      var opq=String(cv[q][cOp]||'').trim(), o2=byOp[opq];
      if(o2 && nkey(o2.paciente)===nkey(cv[q][cPac]) && Math.abs(o2.saldo-num(cv[q][cMonto]))<=0.02) ocupados[opq]=1;
    }

    var props=[];
    for(var i=1;i<cv.length;i++){
      if(String(cv[i][cNota]||'').toLowerCase().indexOf('auto-ingreso')<0) continue;
      var est=String(cv[i][cEst]||'').toLowerCase();
      if(est==='pagado'||est==='cancelado') continue;
      var monto=num(cv[i][cMonto]); if(monto<=0.01) continue;
      var op0=String(cv[i][cOp]||'').trim();
      var pac0=String(cv[i][cPac]||'').trim();
      var conc=String(cv[i][cConc]||'');
      var bdo=byOp[op0];
      // El OP apunta a una venta de ESTA MISMA persona → el folio es CORRECTO, NO
      // es "folio corrido": jamás lo movemos a otra venta (eso sí rompería dinero).
      // Si el saldo no cuadra, es "saldo desincronizado" (otro problema): se reporta.
      if(bdo && nkey(bdo.paciente)===nkey(pac0)){
        if(Math.abs(bdo.saldo-monto)>0.02)
          props.push({ rowNum:i+1, paciente:pac0, monto:monto, concepto:conc, oldOP:op0, newOP:'', quienEsOldOP:bdo.paciente, status:'saldo-desync' });
        continue;
      }
      // Folio corrido REAL: el OP no existe o pertenece a OTRO paciente.
      var cands0=(idxPac[nkey(pac0)]||[]).filter(function(o){ return Math.abs(o.saldo-monto)<=0.02 && !ocupados[o.op] && o.op!==op0; });
      var cands=cands0;
      if(cands.length>1){
        var cn=conc.toLowerCase();
        var strong=cands.filter(function(o){ return o.prods.some(function(pr){ var pl=String(pr).toLowerCase(); return pl.length>3 && (cn.indexOf(pl)>-1 || pl.indexOf(cn)>-1); }); });
        cands = (strong.length===1) ? strong : [];   // 2+ mismos saldos sin desempate claro → revisión humana, no adivinar
      }
      var quienEsOp = bdo ? bdo.paciente : '(no existe en BD)';
      if(cands.length===1){
        props.push({ rowNum:i+1, paciente:pac0, monto:monto, concepto:conc, oldOP:op0, newOP:cands[0].op, quienEsOldOP:quienEsOp, status:'realinear' });
        ocupados[cands[0].op]=1;
      } else {
        props.push({ rowNum:i+1, paciente:pac0, monto:monto, concepto:conc, oldOP:op0, newOP:'', quienEsOldOP:quienEsOp, status: cands0.length? 'ambiguo':'sin-match' });
      }
    }

    if(!apply){
      return { ok:true, preview:true, total:props.length,
               realinear:props.filter(function(p){return p.status==='realinear';}).length,
               ambiguo:props.filter(function(p){return p.status==='ambiguo';}).length,
               sinMatch:props.filter(function(p){return p.status==='sin-match';}).length,
               saldoDesync:props.filter(function(p){return p.status==='saldo-desync';}).length,
               props:props };
    }

    // 3) APLICAR — re-verifica contra los props recién calculados (anti-stale).
    // El lock ya está tomado desde arriba (anti-TOCTOU): read+verify+write consistentes.
    var confirmar = body.aplicar || props.filter(function(p){return p.status==='realinear';});
    var byRow={}; props.forEach(function(p){ byRow[p.rowNum]=p; });
    var hechos=[];
    confirmar.forEach(function(c){
      var rn=c.rowNum, newOP=String(c.newOP||'').trim();
      var p=byRow[rn];
      if(!p || p.status!=='realinear' || !newOP || newOP!==p.newOP) return;   // solo lo verificado en esta corrida
      cc.getRange(rn, cOp+1).setValue(newOP);
      var concAct=String(cc.getRange(rn, cConc+1).getValue()||'');
      if(/saldo pendiente de op/i.test(concAct)) cc.getRange(rn, cConc+1).setValue('Saldo pendiente de OP '+newOP);
      try{ logAudit(actor,'Cobranza','Alinear folio', p.oldOP, 'OP del cargo', p.oldOP+' (era de '+p.quienEsOldOP+')', newOP+' ('+p.paciente+')'); }catch(e){}
      hechos.push({ oldOP:p.oldOP, newOP:newOP, paciente:p.paciente });
    });
    try{ CacheService.getScriptCache().remove('gas_ingresos_v1'); }catch(e){}
    return { ok:true, aplicados:hechos.length, hechos:hechos };

    } finally { if(lock) lock.releaseLock(); }
  }catch(ex){ return { ok:false, error:ex.message }; }
}

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
  // Crédito a favor disponible por paciente — lo consume el botón "Cobrar / Abonar"
  // (ruta auto-ingreso → abonarIngreso). Nunca debe tumbar la vista.
  var credByPac = {};
  try { credByPac = _cobReadCreditos().byPac || {}; } catch (eCr) { credByPac = {}; }
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
    // ¿Cómo se cobra este renglón? Lo decide quién MANDA sobre su MontoCargo:
    //   auto=true  → lo escribe _cobRegistrarSaldoIngreso (Facturado − Pagado de la OP).
    //                Cobrar = abonarIngreso (sube Pagado; el MontoCargo se recomputa solo).
    //                Decrementarlo a mano aquí sería inútil: la próxima edición del ingreso
    //                lo reescribe y el abono se perdería.
    //   auto=false → cargo explícito (saldo inicial / cargo a crédito, con OP o sin ella).
    //                No hay Pagado que subir: el MontoCargo ES el residual → abonarCargo.
    var esAuto = String(cg.nota || '').toLowerCase().indexOf('auto-ingreso') > -1 && !!cg.op;
    out.push({
      origen: 'registro', rowNum: cg.rowNum, op: cg.op || '—',
      auto: esAuto,
      credito: esAuto ? (credByPac[_cobKeyNom(cg.paciente)] || 0) : 0,
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
      estatus: calc.estatus, montoDebe: calc.montoDebe, mesesDebe: calc.mesesDebe, tarifa: calc.tarifa || 0,
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
// nombreOrig: el nombre SIN normalizar, para el match tolerante de abreviados
// (keyBuscar ya viene aplastado por _cobKeyNom y perdió la forma "Melina PA").
function _cobVentasPaciente(keyBuscar, nombreOrig) {
  var pacienteNombre = nombreOrig || keyBuscar;
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
    // Exacto O tolerante (abreviado 2024/2025 "Melina PA" ↔ "Melina Piña Amaros").
    // _ecMismoPaciente vive en analisis.gs; en Apps Script comparten scope global.
    if (_cobKeyNom(row[iPac]) !== keyBuscar &&
        !(typeof _ecMismoPaciente === 'function' && _ecMismoPaciente(row[iPac], pacienteNombre))) continue;
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
/* ═══════════ DESCUENTO POR VOLUMEN (escala por agencia) ═══════════
 * Hoja Descuentos_Agencia: Agencia | Producto | Desde | Hasta | Descuento(%)
 * Varias filas por (agencia, producto) = los escalones. El escalón se elige por la
 * cantidad ACUMULADA DEL MES de ese producto para la agencia. Solo aplica a los
 * productos listados. */
function _cobKeyProd(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function _cobEnsureDescuentos() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_DESCUENTOS);
  if (!sh) {
    sh = ss.insertSheet(COBRANZA_DESCUENTOS);
    sh.getRange(1, 1, 1, 5).setValues([['Agencia', 'Producto', 'Desde', 'Hasta', 'Descuento(%)']]);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    // Semilla REPROVIDA: 3 productos, escala 1-3=0 / 4-8=7 / 9+=12
    var prods = ['Valoración inicial RV', 'Preparación endometrial RV', 'Transferencia de embriones congelados (FET) RV'];
    var rows = [];
    prods.forEach(function (p) {
      rows.push(['REPROVIDA', p, 1, 3, 0]);
      rows.push(['REPROVIDA', p, 4, 8, 7]);
      rows.push(['REPROVIDA', p, 9, '', 12]);
    });
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
    sh.setColumnWidth(2, 320);
  }
  return sh;
}
function _cobLoadDescuentos() {
  var out = {}; // agenciaKey -> { productoKey: [{desde,hasta,pct}] }
  var sh = _cobEnsureDescuentos();
  var raw = sh.getDataRange().getValues();
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    var ag = _cobKeyNom(r[0]), prod = _cobKeyProd(r[1]);
    if (!ag || !prod) continue;
    var desde = _cobNum(r[2]) || 1;
    var hastaRaw = String(r[3] == null ? '' : r[3]).trim();
    var hasta = hastaRaw === '' ? 1e9 : (_cobNum(r[3]) || 1e9);
    var pct = _cobNum(r[4]);
    if (!out[ag]) out[ag] = {};
    if (!out[ag][prod]) out[ag][prod] = [];
    out[ag][prod].push({ desde: desde, hasta: hasta, pct: pct });
  }
  return out;
}
function _cobTierPct(tiers, qty) {
  for (var i = 0; i < tiers.length; i++) { if (qty >= tiers[i].desde && qty <= tiers[i].hasta) return tiers[i].pct; }
  return 0;
}
/* ── CONFIG de descuentos por agencia (panel, sin tocar código) ──────────────
 * Script Property COBRANZA_DESC_CFG = { agencias:[{ agencia, activo,
 *   modo:'directo'|'miembro'        (la condición se cuenta por agencia o por paciente),
 *   conteo:'combinado'|'porProducto'(todos los productos de la lista suman juntos, o cada uno aparte),
 *   base:'mesActual'|'mesAnterior'  (lo del mes define su descuento, o lo del mes pasado aplica a futuro),
 *   productos:[...], escalones:[{desde,hasta,pct}], notas }] }
 * Solo los productos de la lista cuentan para la condición Y reciben el descuento.
 * Migración: si no hay config, se construye desde la hoja legacy Descuentos_Agencia. */
var COBRANZA_DESC_CFG_KEY = 'COBRANZA_DESC_CFG';
function _cobDescCfgMigrar() {
  var out = [];
  try {
    var sh = _cobEnsureDescuentos();
    var raw = sh.getDataRange().getValues();
    var byAg = {}, order = [];
    for (var i = 1; i < raw.length; i++) {
      var ag = String(raw[i][0] || '').trim(); if (!ag) continue;
      var prod = String(raw[i][1] || '').trim();
      var k = _cobKeyNom(ag);
      if (!byAg[k]) { byAg[k] = { agencia: ag, activo: true, modo: 'directo', conteo: 'porProducto', base: 'mesActual', productos: [], _ps: {}, escalones: [], _es: {}, notas: '' }; order.push(k); }
      var A = byAg[k];
      if (prod && !A._ps[_cobKeyProd(prod)]) { A._ps[_cobKeyProd(prod)] = 1; A.productos.push(prod); }
      var desde = _cobNum(raw[i][2]) || 1;
      var hastaRaw = String(raw[i][3] == null ? '' : raw[i][3]).trim();
      var hasta = hastaRaw === '' ? '' : _cobNum(raw[i][3]);
      var pct = _cobNum(raw[i][4]);
      var ek = desde + '|' + hasta + '|' + pct;
      if (!A._es[ek]) { A._es[ek] = 1; A.escalones.push({ desde: desde, hasta: hasta, pct: pct }); }
    }
    order.forEach(function (k) { var A = byAg[k]; delete A._ps; delete A._es; A.escalones.sort(function (a, b) { return a.desde - b.desde; }); out.push(A); });
  } catch (e) {}
  return { agencias: out };
}
function _cobDescCfg() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(COBRANZA_DESC_CFG_KEY);
    if (raw) { var o = JSON.parse(raw); if (o && o.agencias) return o; }
  } catch (e) {}
  var mig = _cobDescCfgMigrar();
  try { PropertiesService.getScriptProperties().setProperty(COBRANZA_DESC_CFG_KEY, JSON.stringify(mig)); } catch (e2) {}
  return mig;
}
function cobDescCfgRead() {
  try { return { ok: true, version: COBRANZA_VER, agencias: (_cobDescCfg().agencias || []) }; }
  catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_VER }; }
}
function cobDescCfgSave(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_egresos'))
      return { ok: false, error: 'Sin permiso para editar los descuentos (requiere editar cuentas por pagar).' };
    var list = [];
    (body.agencias || []).forEach(function (a) {
      var nom = String(a.agencia || '').trim(); if (!nom) return;
      var esc = [];
      (a.escalones || []).forEach(function (t) {
        var h = (t.hasta === '' || t.hasta == null) ? '' : _cobNum(t.hasta);
        esc.push({ desde: _cobNum(t.desde) || 1, hasta: h, pct: _cobNum(t.pct) });
      });
      esc.sort(function (x, y) { return x.desde - y.desde; });
      list.push({
        agencia: nom, activo: a.activo !== false,
        modo: (a.modo === 'miembro' ? 'miembro' : 'directo'),
        conteo: (a.conteo === 'porProducto' ? 'porProducto' : 'combinado'),
        base: (a.base === 'mesAnterior' ? 'mesAnterior' : 'mesActual'),
        productos: (a.productos || []).map(function (p) { return String(p || '').trim(); }).filter(function (p) { return !!p; }),
        escalones: esc, notas: String(a.notas || '')
      });
    });
    PropertiesService.getScriptProperties().setProperty(COBRANZA_DESC_CFG_KEY, JSON.stringify({ agencias: list }));
    try { logAudit(body.usuario || '', 'Cobranza', 'Guardar descuentos agencia', '', '', '', list.length + ' agencia(s)'); } catch (e) {}
    return { ok: true, version: COBRANZA_VER, guardadas: list.length };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_VER }; }
}
function _cobMesPrev(mes) {
  var y = parseInt(String(mes).substring(0, 4), 10), m = parseInt(String(mes).substring(5, 7), 10);
  if (!y || !m) return mes;
  m--; if (m < 1) { m = 12; y--; }
  return y + '-' + ('0' + m).slice(-2);
}
function _cobDescuentosAgencia(agenciaNombre) {
  var agK = _cobKeyNom(agenciaNombre);
  var cfg = null;
  (_cobDescCfg().agencias || []).forEach(function (a) { if (_cobKeyNom(a.agencia) === agK) cfg = a; });
  if (!cfg || cfg.activo === false || !(cfg.productos || []).length || !(cfg.escalones || []).length)
    return { aplica: false, lineas: [], total: 0 };
  var prodSet = {};
  cfg.productos.forEach(function (p) { prodSet[_cobKeyProd(p)] = 1; });
  var tiers = cfg.escalones.map(function (t) { return { desde: _cobNum(t.desde) || 1, hasta: (t.hasta === '' || t.hasta == null) ? 1e9 : (_cobNum(t.hasta) || 1e9), pct: _cobNum(t.pct) }; });
  var porProd = cfg.conteo === 'porProducto', porMiembro = cfg.modo === 'miembro', mesAnt = cfg.base === 'mesAnterior';
  // Partidas elegibles: SOLO los productos de la lista cuentan y reciben descuento.
  var items = [];
  _cobReadCargos().rows.forEach(function (c) {
    if (_cobKeyNom(c.paciente) !== agK) return;
    (c.items || []).forEach(function (it) {
      var pk = _cobKeyProd(it.producto);
      if (!prodSet[pk]) return;
      var f = String(it.fecha || c.fecha || '');
      var mes = f.length >= 7 ? f.substring(0, 7) : '—';
      items.push({ pk: pk, prod: String(it.producto || ''), mes: mes, pac: String(it.pac || '').trim(), cant: _cobNum(it.cantidad) || 1, base: _cobNum(it.total) });
    });
  });
  if (!items.length) return { aplica: false, lineas: [], total: 0 };
  // Conteo del periodo según la regla: [miembro] | [producto] | mes
  function ck(it, mes) { return (porMiembro ? (it.pac || '(sin paciente)') : '') + '|' + (porProd ? it.pk : '') + '|' + mes; }
  var cant = {};
  items.forEach(function (it) { var k = ck(it, it.mes); cant[k] = (cant[k] || 0) + it.cant; });
  // Escalón por partida: si base = mesAnterior, la cantidad del MES PASADO define el % de este mes.
  var out = {};
  items.forEach(function (it) {
    var mesConteo = mesAnt ? _cobMesPrev(it.mes) : it.mes;
    var qty = cant[ck(it, mesConteo)] || 0;
    var pct = _cobTierPct(tiers, qty);
    if (pct <= 0) return;
    var dk = (porMiembro ? (it.pac || '(sin paciente)') : '') + '|' + (porProd ? it.pk : '*') + '|' + it.mes;
    if (!out[dk]) out[dk] = { producto: porProd ? it.prod : 'Productos de la lista', paciente: porMiembro ? (it.pac || '(sin paciente)') : '', mes: it.mes, mesConteo: mesConteo, cantidad: qty, pct: pct, base: 0, descuento: 0 };
    out[dk].base += it.base;
  });
  var lineas = [], total = 0;
  Object.keys(out).forEach(function (k) { var a = out[k]; a.descuento = a.base * a.pct / 100; total += a.descuento; lineas.push(a); });
  lineas.sort(function (x, y) { return x.mes < y.mes ? -1 : x.mes > y.mes ? 1 : ((x.paciente + x.producto) < (y.paciente + y.producto) ? -1 : 1); });
  return { aplica: lineas.length > 0, agencia: String(agenciaNombre).trim(), lineas: lineas, total: total,
           regla: { modo: cfg.modo, conteo: cfg.conteo, base: cfg.base, productos: cfg.productos.length } };
}
function _cobDeposito() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(COBRANZA_DEPOSITO_KEY);
    if (raw) { var o = JSON.parse(raw); return { banco: o.banco || '', beneficiario: o.beneficiario || '', cuenta: o.cuenta || '', clabe: o.clabe || '' }; }
  } catch (e) {}
  return COBRANZA_DEPOSITO_DEF;
}
function readEstadoCobranza(pacienteNombre) {
  try {
    if (!pacienteNombre) return { ok: false, error: 'Nombre de paciente requerido' };
    var keyBuscar = _cobKeyNom(pacienteNombre);

    // Motor A: saldos de este paciente. Exacto O tolerante (abreviado 2024/2025).
    var saldos = _cobBuildSaldos().ops.filter(function (o) {
      return _cobKeyNom(o.paciente) === keyBuscar || o.paciente === pacienteNombre ||
             (typeof _ecMismoPaciente === 'function' && _ecMismoPaciente(o.paciente, pacienteNombre));
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
    var ventas = _cobMasked() ? [] : _cobVentasPaciente(keyBuscar, pacienteNombre);
    var totalVendido = 0, totalPagadoVentas = 0;
    ventas.forEach(function (v) { totalVendido += v.total; totalPagadoVentas += v.pagado; });

    return {
      ok: true, version: COBRANZA_VER,
      paciente: _cobMasked() ? '—' : String(pacienteNombre).trim(),
      masked: _cobMasked(),
      saldos: saldos, totalSaldo: totalSaldo,
      abonos: misAbonos, suscripcion: miSus,
      suscripcionCargos: susCargos, suscripcionCargosTotal: susCargosTotal,
      ventas: ventas, totalVendido: totalVendido, totalPagadoVentas: totalPagadoVentas,
      descuentos: _cobMasked() ? { aplica: false, lineas: [], total: 0 } : _cobDescuentosAgencia(pacienteNombre),
      deposito: _cobDeposito()
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

/* ═════════ REGISTRAR UNA CUENTA POR COBRAR = VENDER A CRÉDITO ═════════
 * "+ Cuenta por cobrar" es SIEMPRE una venta nueva que todavía NO está en Ingresos
 * (decisión del negocio). Antes esta función solo hacía appendRow en Cuentas_Cobrar:
 * creaba la DEUDA pero nunca la VENTA. Al cobrarla el dinero entraba al banco y el P&L
 * jamás veía el ingreso → había que capturarlo a mano, duplicándolo.
 *
 * La raíz estaba en el ALTA, no en el cobro: al cobrar es imposible saber si la venta
 * ya existe; al crear el cargo, sí se sabe (siempre es nueva).
 *
 * Ahora el cargo NACE de la venta:
 *   1) saveIngreso(... pagado:0 ...) → escribe la venta en BD_Ingresos con Pagado 0.
 *      pagado:0 NO manda nada al banco: saveIngreso no llama a saveBankRow ni a
 *      saveCajaChicaIngreso en ninguna rama (el banco lo escribe el FRONT al capturar,
 *      o _abonoRutearABanco al cobrar). Vender a crédito no mueve efectivo: correcto.
 *   2) saveIngreso ve saldo (TotalPagar − Pagado = total) y llama a
 *      _cobRegistrarSaldoIngreso → ESA es la cuenta por cobrar, con OP y Nota
 *      'auto-ingreso' → _cobBuildSaldos la marca auto:true → el botón "Cobrar / Abonar"
 *      rutea a abonarIngreso: sube el Pagado de la venta Y manda el dinero al banco.
 *      UN SOLO MOTOR de cobro. Los abonos parciales salen gratis (Pagado sube de a poco).
 * Esta función YA NO hace appendRow propio: sería una SEGUNDA cuenta por cobrar por el
 * mismo dinero. Son mutuamente excluyentes (return antes del appendRow que ya no existe).
 * La fecha de la venta = la fecha del cargo (devengado: el ingreso pertenece al mes en
 * que se vendió, no a aquel en que se cobre).
 */
/* ── TITULAR: resolución contra el catálogo de Orígenes ────────────────────
   Puro (recibe el catálogo ya leído) → se prueba sin hojas.
   body.titularTipo === 'origen'  → el titular es una agencia / médico externo /
     grupo / coordinador: se resuelve por ID, nombre o alias y se devuelve el
     NOMBRE CANÓNICO del catálogo. Si no resuelve → se rechaza (no se inventa un
     titular nuevo: eso es exactamente lo que rompe el empate del crédito).
   Cualquier otro valor → paciente: texto libre, como siempre (no se rompe el
     flujo normal, que es el 95% de las altas). */
function _cobResolverTitularEn(body, catalogo) {
  body = body || {};
  var raw = String(body.titularId || body.paciente || '').trim();
  if (!raw) return { ok: false, error: 'Falta el titular de la cuenta por cobrar (a quién se le cobra).' };
  if (String(body.titularTipo || '').trim().toLowerCase() !== 'origen')
    return { ok: true, paciente: String(body.paciente || '').trim(), origenId: '' };
  var hit = _origResolverEn(catalogo || [], raw);
  if (!hit || !hit.id)
    return { ok: false, error: 'El titular "' + raw + '" no está en el catálogo de Orígenes externos (agencias / médicos / grupos). ' +
      'Dalo de alta ahí primero y vuelve a elegirlo del menú: si se captura a mano, su crédito a favor queda a un nombre distinto y no se le puede aplicar nunca.' };
  return { ok: true, paciente: hit.nombre, origenId: hit.id, origenTipo: hit.tipo };
}
/* Envoltura que lee el catálogo. Guard de version skew: sin origenes.gs desplegado
   NO se degrada a texto libre en silencio (patrón finance.gs:616-630). */
function _cobResolverTitular(body) {
  if (String((body || {}).titularTipo || '').trim().toLowerCase() !== 'origen')
    return _cobResolverTitularEn(body, []);
  if (typeof _origResolver !== 'function' || typeof readOrigenes !== 'function')
    return { ok: false, error: 'Agrega/actualiza origenes.gs en el proyecto de Apps Script y redespliega: el titular de esta cuenta es una agencia o un médico y debe validarse contra el catálogo de Orígenes.' };
  var reg;
  try { reg = readOrigenes(); } catch (ex) { reg = null; }
  if (!reg || !reg.ok) return { ok: false, error: 'No se pudo leer el catálogo de Orígenes externos para validar el titular: ' + ((reg && reg.error) || 'error desconocido') };
  return _cobResolverTitularEn(body, reg.origenes || []);
}

function cargarSaldoInicial(body) {
  try {
    body = body || {};
    // ── Version skew: si el backend está a medio desplegar, fallar RUIDOSO. Sin
    // saveIngreso esto crearía otra vez deuda sin venta — justo el bug que cierra. ──
    if (typeof saveIngreso !== 'function')
      return { ok: false, error: 'Actualiza finance.gs en Apps Script y redespliega (falta saveIngreso): una cuenta por cobrar ahora se registra como venta a crédito.', version: COBRANZA_VER };
    if (typeof _cobRegistrarSaldoIngreso !== 'function')
      return { ok: false, error: 'Actualiza cobranza.gs en Apps Script y redespliega (falta _cobRegistrarSaldoIngreso).', version: COBRANZA_VER };
    if (typeof INGRESOS_IDS === 'undefined' || typeof INGRESOS_SS_ID === 'undefined')
      return { ok: false, error: 'Actualiza finance.gs en Apps Script y redespliega (faltan los libros de ingresos).', version: COBRANZA_VER };

    // ── TITULAR CANÓNICO ───────────────────────────────────────────────────────
    // El crédito a favor (comisiones.gs) se consume por NOMBRE normalizado contra el
    // titular de la venta. Si el catálogo dice "Dr. Paladino" y aquí se teclea "Juan
    // Paladino", el crédito nace y se queda inerte. Cuando el titular es un origen
    // (agencia/médico/grupo) se RESUELVE contra el catálogo y se guarda su nombre
    // canónico — nunca lo que llegó por el cable.
    var _tit = _cobResolverTitular(body);
    if (!_tit.ok) return { ok: false, error: _tit.error, version: COBRANZA_VER };
    var paciente = _tit.paciente;

    var items = _cobParseItems(body.items);
    var monto = items.length ? _cobItemsMonto(items) : _cobNum(body.monto);
    if (monto <= 0) return { ok: false, error: 'Captura al menos una partida (producto) o un monto.', version: COBRANZA_VER };

    var fecha = body.fecha ? _cobStr(_cobD(body.fecha)) : _cobStr(_cobToday());
    var concepto = String(body.concepto || '').trim();
    if (!concepto) concepto = items.length ? (items.map(function (it) { return it.producto; }).filter(Boolean).slice(0, 3).join(', ') || 'Cuenta por cobrar') : 'Saldo inicial';

    // ── Cargo con MONTO SUELTO (sin partidas): saveIngreso aborta si no hay líneas
    // ('No hay productos en la operación'). Se sintetiza UNA línea desde concepto+monto
    // en vez de exigir partidas: ese flujo funciona hoy y romperlo sería una regresión. ──
    if (!items.length) items = [{ producto: concepto, cantidad: 1, precio: monto, total: monto, saldo: 0, pac: '', fecha: '' }];

    // ── GUARD DEL LIBRO ANUAL ──────────────────────────────────────────────────
    // saveIngreso SIEMPRE escribe en INGRESOS_SS_ID (el libro del año en curso). Un cargo
    // fechado en otro año caería en el libro equivocado con una fecha que no le toca:
    // invisible donde debería estar y contaminando donde no. El guard es contra el LIBRO,
    // no contra el reloj → se auto-mantiene cuando se agregue el libro del año siguiente.
    var anio = parseInt(String(fecha).substring(0, 4), 10);
    if (!anio || INGRESOS_IDS[anio] !== INGRESOS_SS_ID) {
      return { ok: false, version: COBRANZA_VER,
        error: 'La fecha del cargo (' + fecha + ') no corresponde al libro de ingresos activo. Una cuenta por cobrar ahora registra la VENTA en Ingresos, y la venta se escribiría en el libro equivocado. Captura el cargo con una fecha del año en curso, o registra la venta directamente en el libro de ' + anio + '.' };
    }

    // ── Doble alta (doble clic / "Reintentar" tras un timeout que en realidad SÍ guardó):
    // antes duplicaba una deuda (feo pero visible); ahora duplicaría REVENUE. Se rechaza
    // un cargo idéntico (titular + fecha + monto) creado hace segundos. Es una malla, no
    // un candado: dos POST simultáneos pueden leer antes de que el otro escriba. La
    // defensa principal es el front (botón deshabilitado + sin "Reintentar" a ciegas). ──
    var dup = _cobCargoDuplicadoReciente(paciente, fecha, monto, 120);
    if (dup) {
      return { ok: false, version: COBRANZA_VER, duplicado: true, op: dup.op,
        error: 'Ya se registró hace unos segundos una cuenta por cobrar idéntica de ' + paciente + ' por ' + monto.toFixed(2) + ' (venta ' + dup.op + '). No se creó otra para no duplicar el ingreso. Recarga la lista: si de verdad son dos cargos distintos, espera un momento y vuelve a capturarlo.' };
    }

    // ── TITULAR ────────────────────────────────────────────────────────────────
    // El titular de la venta y de la cuenta por cobrar es el paciente del CARGO: en modo
    // agencia (REPROVIDA) cada partida viaja con su propio `pac` (la paciente real), pero
    // a quien se le cobra es a la AGENCIA. Tomar items[0].pac aquí pondría la deuda a
    // nombre de la paciente: el aging por agencia perdería el saldo y se le cobraría a
    // quien no debe. El nombre real de cada partida NO se pierde: viaja en el JSON de
    // Items del renglón (que es lo que se despliega) y resumido en Observaciones.
    var nota = _cobSanitizeNota(body.nota);
    var pacsReales = [];
    for (var pi = 0; pi < items.length; pi++) {
      var pr = String(items[pi].pac || '').trim();
      if (pr && pacsReales.indexOf(pr) < 0) pacsReales.push(pr);
    }
    var opRef = String(body.op || '').trim();   // referencia informativa; la venta genera su propia OP
    var obs = 'Venta a crédito (cuenta por cobrar)'
      + (opRef ? ' · ref. OP ' + opRef : '')
      + (pacsReales.length ? ' · Pacientes: ' + pacsReales.join(', ') : '')
      + (nota ? ' · ' + nota : '');

    // ── Líneas de la venta. El PVP se deriva de total/cantidad para que el total que
    // calcula saveIngreso (pvp × cant) sea EXACTAMENTE el monto del cargo: así la venta
    // y la deuda nacen cuadradas al centavo. ──
    var cat = String(body.categoria || '');
    var lineas = [];
    for (var li = 0; li < items.length; li++) {
      var it = items[li];
      var cant = _cobNum(it.cantidad) || 1;
      var tot = _cobNum(it.total);
      lineas.push({ categoria: cat, producto: String(it.producto || concepto), pvp: (cant ? tot / cant : tot),
                    descuento: 0, cantidad: cant, pagado: 0, ciclo: '' });
    }

    // ── LA VENTA. pagado:0 → _pagadoOpDefinido es true (0 está definido) → saldoGenerado
    // = totalOP → UNA llamada a _cobRegistrarSaldoIngreso → UN renglón en Cuentas_Cobrar.
    // Los campos espejean lo que manda el formulario real de Ingresos para que la fila no
    // ensucie el Estado de Resultados. conciliacion:false porque no hay nada en el banco
    // todavía (marcarla sería mentir); formaPago vacío: aún no se cobra — la forma de pago
    // se define al cobrar, y es ahí donde abonarIngreso rutea el dinero al banco. ──
    var res = saveIngreso({
      token: body.token || '', usuario: String(body.usuario || ''),
      paciente: paciente, fecha: fecha,
      formaPago: '', pagado: 0, pagos: [],
      sucursal: String(body.sucursal || ''), moneda: String(body.moneda || 'MX'), ciclo: '',
      observaciones: obs, factura: '', poliza: '', razonSocial: String(body.razonSocial || ''),
      montoFactMes: 0, facturacion: false, conciliacion: false, contabilidad: false,
      origenExterno: String(body.origenExterno || ''),
      lineas: lineas
    });
    if (!res || !res.ok)
      return { ok: false, version: COBRANZA_VER, error: 'No se pudo registrar la venta de la cuenta por cobrar: ' + ((res && res.error) || 'error desconocido') + '. No se creó ninguna deuda.' };

    // ── El renglón de deuda lo crea _cobRegistrarSaldoIngreso DENTRO de saveIngreso, y ahí
    // los errores se tragan (try/catch vacío: nunca deben tumbar una venta ya escrita). Si
    // no quedó, tendríamos la VENTA sin su DEUDA y esta función devolvería ok:true sobre
    // una cuenta por cobrar que no existe. Se verifica y se falla RUIDOSO. ──
    var sh = _cobEnsureCargos();
    var rowNum = _cobFindAutoRow(sh, res.op);
    if (rowNum < 0) {
      return { ok: false, version: COBRANZA_VER, op: res.op, ventaCreada: true,
        error: 'La VENTA se registró (' + res.op + ') pero NO se pudo crear su cuenta por cobrar. El ingreso YA está capturado: NO lo vuelvas a capturar. Avisa al administrador para que revise la hoja ' + COBRANZA_CARGOS + ' (la deuda de ' + res.op + ' falta).' };
    }

    // Detalle del renglón: el JSON de partidas conserva el paciente real por línea (lo que
    // se despliega en agencias) y el concepto que escribió el usuario.
    try {
      sh.getRange(rowNum, _cobCargosItemsCol(sh)).setValue(JSON.stringify(items));
      var lcC = Math.max(sh.getLastColumn(), 1);
      var hdrC = sh.getRange(1, 1, 1, lcC).getValues()[0].map(function (x) { return String(x).trim().toLowerCase(); });
      var iConcC = hdrC.indexOf('concepto'); if (iConcC < 0) iConcC = 4;
      var iUsrC = hdrC.indexOf('usuario');   if (iUsrC < 0) iUsrC = 8;
      sh.getRange(rowNum, iConcC + 1).setValue(concepto);
      sh.getRange(rowNum, iUsrC + 1).setValue(String(body.usuario || ''));
    } catch (eD) {}

    if (typeof logAudit === 'function') {
      try { logAudit(String(body.usuario || '') || 'sistema', 'Cobranza', 'Registrar cuenta por cobrar (venta a crédito)',
              res.op, 'Venta+CxC', '', paciente + ' · ' + Number(res.total || monto).toFixed(2) + ' · ' + concepto); } catch (ae) {}
    }
    return { ok: true, version: COBRANZA_VER, fecha: fecha, monto: Number(res.total || monto),
             op: res.op, rowNum: rowNum, lineas: res.lineas, paciente: paciente };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}

/* 'auto-ingreso' es un MARCADOR DEL SISTEMA en la columna Nota: significa "este saldo lo
 * manda la venta". Si un POST lo mete en una nota de usuario, el renglón finge venir de
 * una venta que no existe: abonarCargo lo rechaza y _cobBuildSaldos lo marca auto:false
 * (exige OP) → callejón sin salida, imposible de cobrar. Se remueve el token. */
function _cobSanitizeNota(x) {
  return String(x || '').replace(/auto[\s\-_]*ingresos?/gi, '').replace(/\s{2,}/g, ' ').trim();
}
// Fila (1-based) del renglón auto-generado de una OP en Cuentas_Cobrar; -1 si no existe.
function _cobFindAutoRow(sh, op) {
  op = String(op || '').trim(); if (!op) return -1;
  var lr = sh.getLastRow(); if (lr < 2) return -1;
  var lc = Math.max(sh.getLastColumn(), 1);
  var vals = sh.getRange(1, 1, lr, lc).getValues();
  var hdr = (vals[0] || []).map(function (x) { return String(x).trim().toLowerCase(); });
  var iOp = hdr.indexOf('op');     if (iOp < 0) iOp = 1;
  var iNota = hdr.indexOf('nota'); if (iNota < 0) iNota = 7;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][iOp] || '').trim() === op &&
        String(vals[r][iNota] || '').toLowerCase().indexOf('auto-ingreso') > -1) return r + 1;
  }
  return -1;
}
/* Borra el renglón auto-generado de deuda de la OP <op>. Lo llama deleteIngreso
 * (finance.gs) cuando la VENTA se borra por completo.
 *
 * ¿Por qué BORRAR y no cerrar en 0 como hace la cancelación? Porque son cosas distintas:
 *   · CANCELAR conserva la venta como registro histórico (Cancelada=true, con motivo) →
 *     su deuda se cierra y se re-etiqueta 'Cancelado': queda traza de que existió.
 *   · BORRAR elimina la venta del libro: no queda nada de qué ser deuda. Dejar el renglón
 *     (aunque fuera en 0) sería traza de una venta que ya no existe, y en Cuentas por
 *     Cobrar aparecería un adeudo HUÉRFANO con una OP inexistente — justo lo que
 *     auditarCargosSinVenta tiene que salir a cazar después.
 * Sólo toca renglones con Nota 'auto-ingreso' (los que manda la venta). Un cargo
 * capturado a mano NO se borra: no nació de esta venta y no le toca a este flujo.
 * Devuelve { ok, borrado:bool, row } y NUNCA avienta: borrar el ingreso ya ocurrió y no
 * puede tumbarse por esto. */
function _cobBorrarSaldoIngreso(op) {
  try {
    op = String(op || '').trim();
    if (!op) return { ok: false, borrado: false, error: 'OP vacía' };
    var sh = _cobEnsureCargos();
    var rn = _cobFindAutoRow(sh, op);
    if (rn < 2) return { ok: true, borrado: false, row: 0 };   // no había deuda auto: nada que hacer
    sh.deleteRow(rn);
    return { ok: true, borrado: true, row: rn };
  } catch (e) { return { ok: false, borrado: false, error: e.message }; }
}
// ¿Se registró hace <ventanaSeg> una cuenta por cobrar idéntica? (anti doble-alta)
function _cobCargoDuplicadoReciente(paciente, fecha, monto, ventanaSeg) {
  try {
    var sh = _cobEnsureCargos();
    var lr = sh.getLastRow(); if (lr < 2) return null;
    var lc = Math.max(sh.getLastColumn(), 1);
    var hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).trim().toLowerCase(); });
    var iOp = hdr.indexOf('op');           if (iOp < 0) iOp = 1;
    var iPac = hdr.indexOf('paciente');    if (iPac < 0) iPac = 2;
    var iMonto = hdr.indexOf('montocargo');if (iMonto < 0) iMonto = 5;
    var iNota = hdr.indexOf('nota');       if (iNota < 0) iNota = 7;
    var iFecha = hdr.indexOf('fecha');     if (iFecha < 0) iFecha = 0;
    var iTs = hdr.indexOf('timestamp');    if (iTs < 0) iTs = 9;
    var n = Math.min(lr - 1, 60);                  // un duplicado es RECIENTE → solo la cola
    var start = lr - n + 1;
    var vals = sh.getRange(start, 1, n, lc).getValues();
    var ahora = new Date().getTime();
    for (var i = vals.length - 1; i >= 0; i--) {
      var row = vals[i];
      if (String(row[iNota] || '').toLowerCase().indexOf('auto-ingreso') < 0) continue;
      if (_cobKeyNom(row[iPac]) !== _cobKeyNom(paciente)) continue;
      if (Math.abs(_cobNum(row[iMonto]) - _cobNum(monto)) > 0.01) continue;
      if (_cobStr(_cobD(row[iFecha])) !== String(fecha)) continue;
      var ts = (row[iTs] instanceof Date) ? row[iTs].getTime() : 0;
      if (!ts || (ahora - ts) > (ventanaSeg * 1000)) continue;
      return { rowNum: start + i, op: String(row[iOp] || '').trim(), ts: ts };
    }
    return null;
  } catch (e) { return null; }
}

/* ═════════ COBRAR / ABONAR un CARGO EXPLÍCITO de Cuentas_Cobrar ═════════
 * Hermano de finance.abonarIngreso, para los renglones que NO nacen de una venta con
 * pago parcial: saldos iniciales y cargos a crédito (ej. REPROVIDA), con o sin OP.
 * Ahí NO hay Pagado que subir en BD_Ingresos (no hay líneas propias en la venta): el
 * saldo canónico ES el MontoCargo del renglón (Motor A de saldos explícitos), así que
 * abonar = decrementar ese MontoCargo. Atómico (LockService):
 *   1) Cuentas_Cobrar → MontoCargo −= abono. Si llega a ~0 se cierra (Estatus 'Pagado')
 *      y sale de la lista (_cobBuildSaldos salta 'pagado'; _cobReadCargos salta monto 0).
 *   2) Abonos_Cobrar → historial con tipo 'abono-cargo'. Ese tipo NO se resta por OP en
 *      _cobReadAbonos: el saldo YA bajó en el MontoCargo → restarlo otra vez sería doble
 *      conteo (justo lo que pasa con los abonos viejos tipo 'saldo').
 *   3) Conciliación bancaria → UN movimiento de INGRESO por el efectivo, ruteado por forma
 *      de pago con el MISMO mapa que un ingreso (_abonoRutearABanco), etiquetado
 *      [CxC #<fila>] — no [OP #…]: este dinero no pertenece a ninguna OP.
 * Cuentas_Cobrar no tiene columna ID estable → tras tomar el lock se RE-LEE y RE-VALIDA el
 * renglón (paciente, concepto y saldo esperados) antes de escribir: si alguien borró filas
 * entre el render y el guardado, el rowNum apuntaría a OTRA cuenta. Se rechaza en vez de
 * abonarle al renglón equivocado. Ese mismo candado hace que un doble clic no duplique.
 */
function abonarCargo(body) {
  try {
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_ingresos'))
      return { ok: false, error: 'Sin autorización para cobrar/abonar (requiere el permiso editar_ingresos).' };
    // Fail-closed ante deploy parcial: si falta el ruteo a bancos NO se escribe nada
    // (si no, el saldo bajaría y el dinero jamás llegaría a conciliación — el bug original).
    if (typeof _abonoRutearABanco !== 'function')
      return { ok: false, error: 'Actualiza finance.gs en Apps Script y redespliega (falta _abonoRutearABanco).' };
    var rn = parseInt(body.rowNum, 10);
    if (!rn || rn < 2) return { ok: false, error: 'Fila inválida' };
    var monto = _cobNum(body.monto);
    if (monto <= 0.01) return { ok: false, error: 'Nada que abonar (monto en 0).' };
    var formaPago = String(body.formaPago || 'Transferencia').trim();
    var comisionMP = Math.abs(_cobNum(body.comisionMP));
    var fecha = body.fecha ? _cobStr(_cobD(body.fecha)) : _cobStr(_cobToday());
    var usuario = String(body.usuario || '');

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado, reintenta en un momento.' };
    try {
      var sh = SpreadsheetApp.openById(INGRESOS_SS_ID).getSheetByName(COBRANZA_CARGOS);
      if (!sh) return { ok: false, error: 'No se encontró la hoja ' + COBRANZA_CARGOS };
      if (rn > sh.getLastRow()) return { ok: false, error: 'Esa cuenta por cobrar ya no existe. Recarga la lista y reintenta.' };
      var lc = Math.max(sh.getLastColumn(), 1);
      var hdr = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).trim().toLowerCase(); });
      var hi = function (name, def) { var k = hdr.indexOf(name); return k > -1 ? k : def; };
      var iOp = hi('op', 1), iPac = hi('paciente', 2), iConc = hi('concepto', 4),
          iMonto = hi('montocargo', 5), iEst = hi('estatus', 6), iNota = hi('nota', 7);
      var row = sh.getRange(rn, 1, 1, lc).getValues()[0];

      // ── RE-VALIDACIÓN del renglón (sin ID estable, la fila pudo moverse o borrarse) ──
      var saldoActual = _cobNum(row[iMonto]);
      var estActual = _cobLower(row[iEst]);
      if (estActual === 'pagado' || estActual === 'cancelado')
        return { ok: false, error: 'Esa cuenta por cobrar ya está ' + estActual + '. Recarga la lista.' };
      if (String(row[iNota] || '').toLowerCase().indexOf('auto-ingreso') > -1)
        return { ok: false, error: 'Ese saldo viene de la venta ' + String(row[iOp] || '') + ': cóbralo con "Cobrar / Abonar" para que suba el Pagado de la venta.' };
      if (body.paciente != null && String(body.paciente) !== '' && _cobKeyNom(body.paciente) !== _cobKeyNom(row[iPac]))
        return { ok: false, error: 'La cuenta por cobrar cambió (otro paciente). Recarga la lista y reintenta.' };
      if (body.concepto != null && String(body.concepto) !== '' && _cobLower(body.concepto) !== _cobLower(row[iConc]))
        return { ok: false, error: 'La cuenta por cobrar cambió (otro concepto). Recarga la lista y reintenta.' };
      if (body.saldoEsperado != null && String(body.saldoEsperado) !== '' && Math.abs(_cobNum(body.saldoEsperado) - saldoActual) > 0.01)
        return { ok: false, error: 'El saldo cambió (ahora es ' + saldoActual.toFixed(2) + '). Recarga la lista y reintenta.' };
      if (monto > saldoActual + 0.01)
        return { ok: false, error: 'El abono (' + monto.toFixed(2) + ') supera el saldo de la cuenta (' + saldoActual.toFixed(2) + '). Ajusta el monto.' };

      var paciente = String(row[iPac] || '').trim();
      var opRef = String(row[iOp] || '').trim();
      var nuevoSaldo = saldoActual - monto; if (nuevoSaldo < 0) nuevoSaldo = 0;

      // 1) Saldo canónico: decrementa el MontoCargo y cierra el renglón si quedó en 0.
      sh.getRange(rn, iMonto + 1).setValue(nuevoSaldo);
      if (nuevoSaldo <= 0.01) sh.getRange(rn, iEst + 1).setValue('Pagado');

      // 2) Historial del abono (tipo 'abono-cargo' → no se resta otra vez).
      try {
        registrarAbono({ op: opRef, paciente: paciente, monto: monto, tipo: 'abono-cargo',
          formaPago: formaPago, fecha: fecha, usuario: usuario,
          nota: String(body.nota || '') || ('Abono a cuenta por cobrar [CxC #' + rn + ']') });
      } catch (eA) {}

      // 3) Movimiento bancario por el efectivo cobrado, etiquetado [CxC #<fila>].
      var bancoInfo = _abonoRutearABanco('', paciente, monto, formaPago, fecha, comisionMP,
        'Cobro/abono cuenta por cobrar · Px. ' + paciente + ' [CxC #' + rn + ']');

      try { CacheService.getScriptCache().remove('erp_banks_v1'); } catch (e) {}
      if (typeof logAudit === 'function') {
        try { logAudit(usuario || 'sistema', 'Cobranza', 'Abonar cuenta por cobrar', 'CxC #' + rn, 'MontoCargo',
                saldoActual.toFixed(2), nuevoSaldo.toFixed(2) + ' (abono ' + monto.toFixed(2) + ' · ' + formaPago + ')'); } catch (ae) {}
      }

      return { ok: true, version: COBRANZA_VER, rowNum: rn, paciente: paciente, op: opRef,
               abonado: monto, saldoAntes: saldoActual, saldo: nuevoSaldo,
               cerrado: (nuevoSaldo <= 0.01), formaPago: formaPago, banco: bancoInfo };
    } finally { lock.releaseLock(); }
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_VER };
  }
}

/* ═════════ AUDITORÍA: cargos SIN VENTA (SOLO LECTURA) ═════════
 * Correr A MANO desde el editor de Apps Script. NO repara nada, NO escribe nada, NO tiene
 * botón: solo LISTA. Cada caso se decide a mano.
 *
 * Qué lista: las cuentas por cobrar SIN la nota 'auto-ingreso', o sea las que se crearon
 * ANTES del fix, cuando "+ Cuenta por cobrar" solo escribía la DEUDA y nunca la VENTA. En
 * esas, cuando se cobra, el dinero entra al banco pero el ingreso no existe en el P&L.
 * (Las nuevas nacen de la venta y traen 'auto-ingreso': esas ya están cuadradas.)
 *
 * ⚠ LA COLUMNA "PISTA" NO ES UNA PRUEBA. Solo dice si ese paciente tiene ALGO en
 * BD_Ingresos, empatando ÚNICAMENTE POR NOMBRE — sin OP, sin fecha, sin monto. Un "sí"
 * NO significa que la venta de ESTE cargo esté registrada: puede ser cualquier otra venta
 * del mismo paciente (justo el falso indicio que casi nos hace duplicar un ingreso). Un
 * "no" sí es señal fuerte de que la venta falta. Verifica caso por caso antes de tocar nada.
 *
 * Los cargos ya PAGADOS también salen: si se cobraron sin venta, el dinero entró al banco
 * y el ingreso nunca se registró — el descuadre ya ocurrió y sigue ahí.
 */
function auditarCargosSinVenta() {
  var L = [];
  function log(s) { L.push(s); try { Logger.log(s); } catch (e) {} }
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var sh = ss.getSheetByName(COBRANZA_CARGOS);
    if (!sh) return 'No existe la hoja ' + COBRANZA_CARGOS + ' — nada que auditar.';
    var lr = sh.getLastRow(); if (lr < 2) return 'La hoja ' + COBRANZA_CARGOS + ' está vacía — nada que auditar.';
    var lc = Math.max(sh.getLastColumn(), 1);
    var vals = sh.getRange(1, 1, lr, lc).getValues();
    var hdr = (vals[0] || []).map(function (x) { return String(x).trim().toLowerCase(); });
    var iFecha = hdr.indexOf('fecha');      if (iFecha < 0) iFecha = 0;
    var iOp = hdr.indexOf('op');            if (iOp < 0) iOp = 1;
    var iPac = hdr.indexOf('paciente');     if (iPac < 0) iPac = 2;
    var iCat = hdr.indexOf('categoria');    if (iCat < 0) iCat = 3;
    var iConc = hdr.indexOf('concepto');    if (iConc < 0) iConc = 4;
    var iMonto = hdr.indexOf('montocargo'); if (iMonto < 0) iMonto = 5;
    var iEst = hdr.indexOf('estatus');      if (iEst < 0) iEst = 6;
    var iNota = hdr.indexOf('nota');        if (iNota < 0) iNota = 7;

    // Ventas por paciente (SOLO por nombre → es una PISTA, no una prueba).
    var ventasPorPac = {};
    try {
      var shIn = ss.getSheetByName(BD_INGRESOS_TAB);
      if (shIn) {
        var din = shIn.getDataRange().getValues();
        var HI = (din[0] || []).map(function (x) { return _cobLower(x); });
        var hcI = function () { for (var a = 0; a < arguments.length; a++) { var k = HI.indexOf(arguments[a]); if (k > -1) return k; } return -1; };
        var jOp = hcI('op'); if (jOp < 0) jOp = 0;
        var jPac = hcI('paciente'), jTot = hcI('totalpagar', 'total a pagar', 'total'), jFec = hcI('fecha');
        for (var r = 1; r < din.length; r++) {
          var pk = _cobKeyNom(din[r][jPac]); if (!pk) continue;
          if (!ventasPorPac[pk]) ventasPorPac[pk] = { n: 0, total: 0, ops: {} };
          ventasPorPac[pk].total += _cobNum(din[r][jTot]);
          var oo = String(din[r][jOp] || '').trim();
          if (oo && !ventasPorPac[pk].ops[oo]) { ventasPorPac[pk].ops[oo] = _cobStr(_cobD(din[r][jFec])); ventasPorPac[pk].n++; }
        }
      }
    } catch (eI) { log('! No se pudo leer ' + BD_INGRESOS_TAB + ': ' + eI.message + ' (la columna PISTA saldrá vacía)'); }

    var abonos = { byOp: {} };
    try { abonos = _cobReadAbonos(); } catch (eA) { log('! No se pudieron leer los abonos: ' + eA.message); }

    var casos = [], totMonto = 0, totSaldo = 0, nAuto = 0, nPagados = 0;
    for (var q = 1; q < vals.length; q++) {
      var row = vals[q];
      var nota = String(row[iNota] || '').toLowerCase();
      var op = String(row[iOp] || '').trim();
      if (nota.indexOf('auto-ingreso') > -1 && op) { nAuto++; continue; }   // nace de una venta: cuadrado
      var monto = _cobNum(row[iMonto]);
      var pac = String(row[iPac] || '').trim();
      if (!pac && monto <= 0) continue;                                     // fila basura/vacía
      var est = String(row[iEst] || '').trim();
      var abo = op ? (abonos.byOp[op] || 0) : 0;
      var saldo = monto - abo;
      var pk2 = _cobKeyNom(pac);
      var v = ventasPorPac[pk2];
      var pista = v ? ('SÍ tiene ' + v.n + ' venta(s) por ' + v.total.toFixed(2) + ' — NO prueba que ESTA esté registrada')
                    : 'NO hay NINGUNA venta a ese nombre — señal fuerte de que la venta falta';
      if (_cobLower(est) === 'pagado') nPagados++;
      totMonto += monto; totSaldo += Math.max(0, saldo);
      casos.push({ fila: q + 1, fecha: _cobStr(_cobD(row[iFecha])), paciente: pac, op: op || '—',
                   categoria: String(row[iCat] || ''), concepto: String(row[iConc] || ''),
                   monto: monto, abonado: abo, saldo: saldo, estatus: est || 'Pendiente', pista: pista });
    }

    log('══════════ CARGOS SIN VENTA EN INGRESOS (SOLO LECTURA) ══════════');
    log('Hoja: ' + COBRANZA_CARGOS + ' · renglones revisados: ' + (lr - 1));
    log('Nacidos de una venta (nota auto-ingreso, ya cuadrados, NO se listan): ' + nAuto);
    log('SIN venta (creados antes del fix): ' + casos.length + '  ·  de esos, ya cobrados/cerrados: ' + nPagados);
    log('Monto total de esos cargos: ' + totMonto.toFixed(2) + '  ·  saldo vivo: ' + totSaldo.toFixed(2));
    log('');
    log('⚠ La PISTA empata SOLO POR NOMBRE (sin OP, sin fecha, sin monto). Un "SÍ" NO prueba');
    log('  que la venta de ese cargo exista: puede ser otra venta del mismo paciente. Revisa');
    log('  caso por caso antes de tocar nada. Esta función no repara nada.');
    log('');
    for (var c = 0; c < casos.length; c++) {
      var k = casos[c];
      log('── Fila ' + k.fila + ' · ' + k.fecha + ' · ' + k.paciente + ' · OP ' + k.op);
      log('   ' + k.concepto + (k.categoria ? ' [' + k.categoria + ']' : '') + ' · ' + k.estatus);
      log('   Monto ' + k.monto.toFixed(2) + ' · abonado ' + k.abonado.toFixed(2) + ' · saldo ' + k.saldo.toFixed(2));
      log('   PISTA: ' + k.pista);
    }
    if (!casos.length) log('✓ No hay cargos sin venta: todas las cuentas por cobrar nacen de una venta registrada.');
    log('');
    log('Qué hacer con cada uno (lo decides tú, caso por caso):');
    log(' · Si la venta NO está en Ingresos → captúrala ahí (con lo pagado que corresponda) y borra este cargo,');
    log('   o cóbralo y captura el ingreso a mano UNA vez. Lo que no debe pasar es cobrarlo y no registrar nada.');
    log(' · Si la venta SÍ está (verificado a mano, no por la pista) → este cargo es una deuda DUPLICADA: bórralo.');
    log('══════════════════════════════════════════════════════════════════');
    return L.join('\n');
  } catch (ex) {
    log('ERROR en la auditoría: ' + ex.message);
    return L.join('\n');
  }
}

/* ═════════ AUDITORÍA: abonos huérfanos del flujo viejo (SOLO LECTURA) ═════════
 * Correr A MANO desde el editor de Apps Script. NO repara nada, NO escribe nada:
 * solo reporta qué quedó descuadrado por el bug del botón "Abono" de Cuentas por Cobrar.
 *
 * El bug: ese botón llamaba a registrarAbono(), que únicamente hacía appendRow en
 * Abonos_Cobrar con tipo 'saldo'. NUNCA subía el Pagado de BD_Ingresos ni creaba el
 * movimiento bancario. O sea: el abono quedó anotado, pero el dinero no entró a ningún
 * libro. Todo abono con tipo 'saldo' es sospechoso y sale listado aquí.
 *
 * Efecto colateral que hay que entender al leer el reporte:
 *   · Abono tipo 'saldo' CON OP  → _cobReadAbonos SÍ lo resta por OP, así que el saldo en
 *     pantalla SÍ bajó… aunque el Pagado nunca subió y el banco nunca lo vio. El adeudo
 *     se ve saldado y el dinero no está en ninguna cuenta: descuadre silencioso.
 *   · Abono tipo 'saldo' SIN OP  → nunca se restó de nada (_cobBuildSaldos fuerza abo=0
 *     cuando el cargo no trae OP), así que además el saldo jamás bajó: el paciente sigue
 *     apareciendo debiendo lo mismo por más abonos que se le hayan registrado.
 *
 * "Banco" es una BÚSQUEDA HEURÍSTICA por texto (la OP en las Observaciones de Santander /
 * Mercado Pago / AMEX / Caja Chica). Un "sí" no prueba que ese movimiento sea este abono
 * (pudo ser el cobro original de la venta); un "no" sí es señal fuerte de que faltó.
 * Verifica a mano antes de tocar nada.
 */
function auditarAbonosHuerfanos() {
  var L = [];
  function log(s) { L.push(s); try { Logger.log(s); } catch (e) {} }
  try {
    var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
    var shAb = ss.getSheetByName(COBRANZA_ABONOS);
    if (!shAb) return 'No existe la hoja ' + COBRANZA_ABONOS + ' — nada que auditar.';

    // ── OPs de BD_Ingresos: Facturado / Pagado por OP ──
    var opInfo = {};
    try {
      var shIn = ss.getSheetByName(BD_INGRESOS_TAB);
      if (shIn) {
        var din = shIn.getDataRange().getValues();
        var HI = din[0].map(function (x) { return _cobLower(x); });
        var hcI = function () { for (var a = 0; a < arguments.length; a++) { var k = HI.indexOf(arguments[a]); if (k > -1) return k; } return -1; };
        var jOp = hcI('op'); if (jOp < 0) jOp = 0;
        var jTot = hcI('totalpagar', 'total a pagar', 'total'), jPag = hcI('pagado');
        for (var r = 1; r < din.length; r++) {
          var o = String(din[r][jOp] || '').trim(); if (!o) continue;
          if (!opInfo[o]) opInfo[o] = { total: 0, pagado: 0 };
          opInfo[o].total += _cobNum(din[r][jTot]);
          opInfo[o].pagado += (jPag > -1 ? _cobNum(din[r][jPag]) : 0);
        }
      }
    } catch (eI) { log('! No se pudo leer BD_Ingresos: ' + eI.message); }

    // ── Texto de las observaciones de todos los bancos (búsqueda heurística) ──
    var bankTxt = [];
    try {
      if (typeof _bankSheetByKey === 'function' && typeof _bankObsIdx === 'function') {
        var keys = ['santander', 'mercadopago', 'amex'];
        for (var b = 0; b < keys.length; b++) {
          var shB = null;
          try { shB = _bankSheetByKey(keys[b]); } catch (eB0) { shB = null; }
          if (!shB) continue;
          var lrB = shB.getLastRow(); if (lrB < 2) continue;
          var idx = _bankObsIdx(keys[b]);
          var valsB = shB.getRange(2, 1, lrB - 1, shB.getLastColumn()).getValues();
          for (var vb = 0; vb < valsB.length; vb++) bankTxt.push(String(valsB[vb][idx] || ''));
        }
      }
      if (typeof getCajaChicaSheet === 'function') {
        var shC = getCajaChicaSheet();
        if (shC && shC.getLastRow() > 1) {
          var valsC = shC.getRange(2, 1, shC.getLastRow() - 1, shC.getLastColumn()).getValues();
          for (var vc = 0; vc < valsC.length; vc++) {
            for (var cc = 0; cc < valsC[vc].length; cc++) {
              var cel = String(valsC[vc][cc] || ''); if (cel.length > 3) bankTxt.push(cel);
            }
          }
        }
      }
    } catch (eB) { log('! No se pudieron leer los bancos: ' + eB.message); }
    var bankBlob = bankTxt.join('  ');

    // ── Barrido de Abonos_Cobrar ──
    var raw = shAb.getDataRange().getValues();
    var nTot = 0, nSaldo = 0, sumSaldo = 0, sumSinBanco = 0, nSinBanco = 0, nSinOp = 0, sumSinOp = 0;
    var det = [];
    for (var i = 1; i < raw.length; i++) {
      var rr = raw[i];
      var monto = _cobNum(rr[3]); if (monto <= 0) continue;
      nTot++;
      var tipo = _cobLower(rr[6]);
      if (tipo !== 'saldo') continue;              // 'abono-op' / 'abono-cargo' / 'suscripcion' = flujos sanos
      nSaldo++; sumSaldo += monto;
      var op = String(rr[1] || '').trim();
      var pac = String(rr[2] || '').trim();
      var fec = _cobStr(_cobD(rr[0]));
      var enBanco = false;
      if (op) enBanco = bankBlob.indexOf(op) > -1;
      else if (pac) enBanco = bankBlob.indexOf(pac) > -1;
      if (!enBanco) { nSinBanco++; sumSinBanco += monto; }
      var oi = op ? opInfo[op] : null;
      var pagTxt;
      if (!op) { pagTxt = 'sin OP'; nSinOp++; sumSinOp += monto; }
      else if (!oi) pagTxt = 'OP no está en BD_Ingresos';
      else pagTxt = 'Facturado ' + oi.total.toFixed(2) + ' / Pagado ' + oi.pagado.toFixed(2) +
                    ' / residual ' + Math.max(0, oi.total - oi.pagado).toFixed(2);
      det.push('  fila ' + (i + 1) + ' · ' + fec + ' · ' + (op || '(sin OP)') + ' · ' + (pac || '(sin paciente)') +
               ' · $' + monto.toFixed(2) + ' · ' + pagTxt + ' · banco: ' + (enBanco ? 'quizá sí' : 'NO ENCONTRADO'));
    }

    log('═══ AUDITORÍA DE ABONOS HUÉRFANOS (solo lectura) ═══');
    log('Abonos en ' + COBRANZA_ABONOS + ': ' + nTot + ' con monto > 0.');
    log('Del flujo VIEJO (tipo "saldo"): ' + nSaldo + ' por $' + sumSaldo.toFixed(2) + '.');
    log('  · Sin movimiento bancario localizable: ' + nSinBanco + ' por $' + sumSinBanco.toFixed(2) + '.');
    log('  · Sin OP (su saldo NUNCA bajó, siguen apareciendo como deuda): ' + nSinOp + ' por $' + sumSinOp.toFixed(2) + '.');
    log('  · Con OP: el saldo en pantalla SÍ bajó (se resta en _cobReadAbonos), pero el Pagado');
    log('    de BD_Ingresos no subió y el banco no vio el dinero → descuadre silencioso.');
    log('Los tipos "abono-op" y "abono-cargo" son de los flujos nuevos y NO se listan.');
    log(nSaldo ? '─── Detalle ───' : 'Sin abonos del flujo viejo: nada que reconciliar.');
    for (var d = 0; d < det.length; d++) log(det[d]);
    log('═══ Fin. Nada se modificó. Verifica a mano antes de corregir. ═══');
    return L.join('\n');
  } catch (ex) {
    return 'Error en la auditoría: ' + ex.message + '\n' + L.join('\n');
  }
}

/* ═════════════ CRÉDITO A FAVOR DEL PACIENTE (saldo a favor) ═════════════
 * Cuando un ingreso se edita y queda Pagado > Facturado (el paciente pagó de más,
 * p. ej. se le cambió a un producto más barato), el excedente NO se devuelve: se
 * guarda como CRÉDITO a favor del paciente en la hoja Creditos_Favor y queda
 * aplicable a otra (o la misma) cuenta desde el botón "Cobrar / Abonar".
 *
 * Se mantiene en hoja aparte a propósito: NO va en Cuentas_Cobrar para que los
 * motores de deuda (_cobBuildSaldos / readEstadoCuentaPaciente) jamás lo lean como
 * un adeudo (sería un cargo positivo = deuda falsa). Es idempotente por OP: editar
 * el ingreso reescribe el MISMO renglón (upsert), no apila créditos.
 * Esquema Creditos_Favor: Fecha | OP | Paciente | MontoCredito | Nota | Usuario | Timestamp
 */
var COBRANZA_CREDITOS = (typeof COBRANZA_CREDITOS !== 'undefined') ? COBRANZA_CREDITOS : 'Creditos_Favor';
function _cobEnsureCreditos() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_CREDITOS);
  if (!sh) {
    sh = ss.insertSheet(COBRANZA_CREDITOS);
    sh.appendRow(['Fecha', 'OP', 'Paciente', 'MontoCredito', 'Nota', 'Usuario', 'Timestamp']);
    sh.setFrozenRows(1);
  }
  return sh;
}
// Upsert por OP: fija el crédito a favor de esa OP en 'monto' (absoluto). monto<=0
// cierra el crédito (deja el renglón en 0 → sale de "disponible"). La llama
// updateIngreso al detectar sobrepago (o al revertir el sobrepago). Nunca tumba nada.
function _cobRegistrarCreditoFavor(op, paciente, monto, fecha) {
  try {
    op = String(op || '').trim();
    var cred = _cobNum(monto); if (cred < 0) cred = 0;
    var sh = _cobEnsureCreditos();
    var raw = sh.getDataRange().getValues();
    var foundRow = -1;
    if (op) { for (var r = 1; r < raw.length; r++) { if (String(raw[r][1] || '').trim() === op) { foundRow = r + 1; break; } } }
    if (foundRow > 0) {
      sh.getRange(foundRow, 4).setValue(cred);
      if (paciente) sh.getRange(foundRow, 3).setValue(String(paciente).trim());
      sh.getRange(foundRow, 5).setValue(cred > 0.01 ? 'credito-favor' : 'credito-favor-cerrado');
      return;
    }
    if (cred > 0.01) {
      sh.appendRow([ fecha ? _cobStr(_cobD(fecha)) : _cobStr(_cobToday()), op, String(paciente || '').trim(), cred, 'credito-favor', '', new Date() ]);
    }
  } catch (e) {}
}
function _cobReadCreditos() {
  var out = { byPac: {}, rows: [] };
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_CREDITOS);
  if (!sh) return out;
  var raw = sh.getDataRange().getValues();
  for (var i = 1; i < raw.length; i++) {
    var monto = _cobNum(raw[i][3]); if (monto <= 0.01) continue;
    var rec = { rowNum: i + 1, fecha: _cobStr(_cobD(raw[i][0])), op: String(raw[i][1] || '').trim(), paciente: String(raw[i][2] || ''), monto: monto };
    out.rows.push(rec);
    var kp = _cobKeyNom(raw[i][2]);
    out.byPac[kp] = (out.byPac[kp] || 0) + monto;
  }
  return out;
}
// Crédito a favor disponible de un paciente → { total, rows }.
function _cobCreditoFavorPaciente(paciente) {
  var all = _cobReadCreditos(); var kp = _cobKeyNom(paciente);
  var rows = all.rows.filter(function (x) { return _cobKeyNom(x.paciente) === kp; });
  var total = 0; rows.forEach(function (x) { total += x.monto; });
  return { total: total, rows: rows };
}
// Consume hasta 'monto' del crédito a favor del paciente (FIFO por renglón),
// decrementando la hoja. Devuelve el monto realmente aplicado (offset interno; NO
// mueve dinero real). La llama abonarIngreso cuando el usuario aplica su crédito.
function _cobAplicarCreditoFavor(paciente, monto) {
  var restante = _cobNum(monto); if (restante <= 0) return 0;
  var sh = _cobEnsureCreditos();
  var raw = sh.getDataRange().getValues();
  var kp = _cobKeyNom(paciente); var aplicado = 0;
  for (var i = 1; i < raw.length && restante > 0.01; i++) {
    if (_cobKeyNom(raw[i][2]) !== kp) continue;
    var disp = _cobNum(raw[i][3]); if (disp <= 0.01) continue;
    var take = Math.min(disp, restante);
    sh.getRange(i + 1, 4).setValue(disp - take);
    if (disp - take <= 0.01) sh.getRange(i + 1, 5).setValue('credito-favor-cerrado');
    aplicado += take; restante -= take;
  }
  return aplicado;
}

/* ═════════════ CONSUMO DE CRÉDITO POR "NOTA DE CRÉDITO" (ledger por OP) ═════════════
 * 'Nota de Crédito' es una forma de pago del alta de ingresos: el paciente paga con
 * el crédito a favor que ya tenía. Antes NO se descontaba de Creditos_Favor → el
 * MISMO crédito se podía gastar infinitas veces. Este ledger arregla eso y hace el
 * consumo IDEMPOTENTE por OP (indispensable: al editar una OP el backend reescribe
 * las líneas y volvería a consumir el crédito de cero).
 *
 * Esquema Creditos_Consumo: Fecha | OP | Paciente | MontoNC | Aplicado | Excedente |
 *                           Autorizado | Usuario | Timestamp
 *   MontoNC   = NC declarada hoy en la OP (absoluto)
 *   Aplicado  = cuánto se decrementó de verdad de Creditos_Favor por esta OP (acumulado)
 *   Excedente = NC que rebasó el crédito disponible (solo >0 con autorización explícita)
 * Va en hoja aparte para que los motores de deuda jamás lo lean como cargo.
 */
var COBRANZA_CONSUMO_NC = (typeof COBRANZA_CONSUMO_NC !== 'undefined') ? COBRANZA_CONSUMO_NC : 'Creditos_Consumo';
function _cobEnsureConsumoNC() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_CONSUMO_NC);
  if (!sh) {
    sh = ss.insertSheet(COBRANZA_CONSUMO_NC);
    sh.appendRow(['Fecha', 'OP', 'Paciente', 'MontoNC', 'Aplicado', 'Excedente', 'Autorizado', 'Usuario', 'Timestamp']);
    sh.setFrozenRows(1);
  }
  return sh;
}
// Lo ya consumido por una OP → { rowNum, titular, montoNC, aplicado, excedente } (ceros si no existe).
// 'titular' = de QUIÉN se consumió el crédito (paciente o médico) — se respeta al editar.
function _cobConsumoNCPorOP(op) {
  var out = { rowNum: -1, titular: '', montoNC: 0, aplicado: 0, excedente: 0 };
  op = String(op || '').trim(); if (!op) return out;
  var sh = _cobEnsureConsumoNC();
  var raw = sh.getDataRange().getValues();
  for (var i = 1; i < raw.length; i++) {
    if (String(raw[i][1] || '').trim() !== op) continue;
    out.rowNum = i + 1; out.titular = String(raw[i][2] || ''); out.montoNC = _cobNum(raw[i][3]);
    out.aplicado = _cobNum(raw[i][4]); out.excedente = _cobNum(raw[i][5]);
    return out;
  }
  return out;
}
// ¿Esta OP ya consumió crédito por NC? (cuenta también el excedente autorizado:
// una OP puede tener aplicado=0 y excedente>0 si el paciente no tenía crédito).
function _cobOPTieneConsumoNC(op) {
  var c = _cobConsumoNCPorOP(op);
  return (c.aplicado + c.excedente) > 0.01;
}
// Upsert del renglón del ledger de esta OP (montoNC/aplicado/excedente ABSOLUTOS).
function _cobRegistrarConsumoNC(op, paciente, montoNC, aplicado, excedente, autorizadoPor, usuario, fecha) {
  var sh = _cobEnsureConsumoNC();
  var prev = _cobConsumoNCPorOP(op);
  var vals = [
    fecha ? _cobStr(_cobD(fecha)) : _cobStr(_cobToday()), String(op || '').trim(), String(paciente || '').trim(),
    _cobNum(montoNC), _cobNum(aplicado), _cobNum(excedente),
    String(autorizadoPor || ''), String(usuario || ''), new Date()
  ];
  if (prev.rowNum > 0) sh.getRange(prev.rowNum, 1, 1, vals.length).setValues([vals]);
  else sh.appendRow(vals);
}
/* Devuelve crédito a favor al paciente cuando una edición BAJA la NC de una OP.
   No intenta "des-decrementar" los renglones FIFO originales (imposible saber
   cuáles fueron): repone el monto como un renglón de crédito propio del paciente,
   económicamente equivalente. Llave OP+':nc-rev' → nunca choca con el upsert por
   OP de _cobRegistrarCreditoFavor, y es ACUMULATIVO (varias bajas suman). */
function _cobRestituirCreditoFavor(op, paciente, monto) {
  var add = _cobNum(monto); if (add <= 0.01) return 0;
  var key = String(op || '').trim() + ':nc-rev';
  var sh = _cobEnsureCreditos();
  var raw = sh.getDataRange().getValues();
  for (var r = 1; r < raw.length; r++) {
    if (String(raw[r][1] || '').trim() !== key) continue;
    sh.getRange(r + 1, 4).setValue(_cobNum(raw[r][3]) + add);
    sh.getRange(r + 1, 5).setValue('credito-favor-nc-reverso');
    return add;
  }
  sh.appendRow([_cobStr(_cobToday()), key, String(paciente || '').trim(), add, 'credito-favor-nc-reverso', '', new Date()]);
  return add;
}

/* ── Aplicación NETA del consumo de NC de una OP (la usan saveIngreso y
   updateIngresoConBancos). Idempotente: corre N veces con la misma NC → un solo
   consumo. Devuelve { ok, aplicado, excedente, disponible, error }.
   'validarSolo' = no escribe nada, solo dice si pasaría (para bloquear ANTES de
   guardar las filas del ingreso).
   El crédito SIEMPRE es del MISMO paciente: _cobCreditoFavorPaciente y
   _cobAplicarCreditoFavor filtran por _cobKeyNom(paciente). ── */
function _cobAplicarNCIngreso(op, paciente, montoNC, opts) {
  opts = opts || {};
  var nc = _cobNum(montoNC); if (nc < 0) nc = 0;
  var prev = _cobConsumoNCPorOP(op);
  // TITULAR EFECTIVO del crédito de esta OP: si ya se registró un consumo, se RESPETA el
  // titular guardado (paciente o médico). Así, al editar una venta cuya NC salió del saldo
  // de un MÉDICO, la devolución/ajuste va al MÉDICO — nunca al paciente por error (evita
  // regalar un crédito a quien no le corresponde). En una OP nueva es el titular recibido.
  var titular = (prev.rowNum > 0 && prev.titular) ? prev.titular : String(paciente || '').trim();
  var delta = nc - prev.aplicado - prev.excedente;   // lo que falta consumir (o devolver)

  if (Math.abs(delta) <= 0.01) return { ok: true, aplicado: prev.aplicado, excedente: prev.excedente, disponible: _cobCreditoFavorPaciente(titular).total, sinCambio: true };

  // Baja de NC → devolver crédito al titular (de quien salió).
  if (delta < 0) {
    if (opts.validarSolo) return { ok: true, aplicado: prev.aplicado, excedente: prev.excedente, disponible: _cobCreditoFavorPaciente(titular).total, devolver: -delta };
    var devolver = -delta;
    // Primero se cancela el excedente autorizado (nunca salió de un crédito real),
    // y solo el resto se repone como crédito del titular.
    var quitaExc = Math.min(prev.excedente, devolver);
    var quitaApl = devolver - quitaExc;
    if (quitaApl > 0.01) _cobRestituirCreditoFavor(op, titular, quitaApl);
    _cobRegistrarConsumoNC(op, titular, nc, prev.aplicado - quitaApl, prev.excedente - quitaExc, '', opts.usuario, opts.fecha);
    return { ok: true, aplicado: prev.aplicado - quitaApl, excedente: prev.excedente - quitaExc, devuelto: quitaApl, disponible: _cobCreditoFavorPaciente(titular).total };
  }

  // Alza de NC → consumir 'delta' del crédito disponible del titular.
  var disp = _cobCreditoFavorPaciente(titular).total;
  if (delta > disp + 0.01) {
    if (!opts.autorizarExcedente) {
      return { ok: false, disponible: disp, requerido: delta,
        error: '"' + String(titular || '').trim() + '" solo tiene $' + disp.toFixed(2) +
               ' de crédito a favor y esta Nota de Crédito requiere $' + delta.toFixed(2) + '. Faltan $' + (delta - disp).toFixed(2) +
               '. Corrige el monto o pide a un autorizado que lo apruebe (permiso autorizar_credito_excedido).' };
    }
    if (opts.validarSolo) return { ok: true, aplicado: prev.aplicado + disp, excedente: prev.excedente + (delta - disp), disponible: disp, autorizado: true };
    var apl = _cobAplicarCreditoFavor(titular, disp);
    var exc = delta - apl;
    _cobRegistrarConsumoNC(op, titular, nc, prev.aplicado + apl, prev.excedente + exc, opts.autorizadoPor || opts.usuario || '', opts.usuario, opts.fecha);
    return { ok: true, aplicado: prev.aplicado + apl, excedente: prev.excedente + exc, disponible: 0, autorizado: true, excedido: exc };
  }

  if (opts.validarSolo) return { ok: true, aplicado: prev.aplicado + delta, excedente: prev.excedente, disponible: disp };
  var aplicado = _cobAplicarCreditoFavor(titular, delta);
  var faltante = delta - aplicado;   // carrera/redondeo: lo no aplicado queda como excedente
  _cobRegistrarConsumoNC(op, titular, nc, prev.aplicado + aplicado, prev.excedente + faltante, faltante > 0.01 ? (opts.autorizadoPor || opts.usuario || '') : '', opts.usuario, opts.fecha);
  return { ok: true, aplicado: prev.aplicado + aplicado, excedente: prev.excedente + faltante, disponible: Math.max(0, disp - aplicado) };
}

/* ═════════════ AUDITORÍA DE NOTAS DE CRÉDITO (solo lectura) ═════════════
 * Corre esta función DESDE EL EDITOR de Apps Script (Ejecutar → auditarNotasCredito)
 * y lee el reporte en el Log (Ver → Registro de ejecución).
 *
 * Contexto: hasta el build que introdujo _cobAplicarNCIngreso, pagar con
 * 'Nota de Crédito' NUNCA decrementaba Creditos_Favor → el mismo crédito se pudo
 * gastar varias veces. Esta función NO corrige nada: solo compara, por paciente,
 * el crédito generado contra la NC realmente usada en BD_Ingresos (todos los años)
 * y marca a quién se le gastó crédito que no tenía.
 */
function auditarNotasCredito() {
  try {
    var FP_NC = 'nota de credito';
    function _normFP(s) {
      return String(s || '').trim().toLowerCase()
        .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
        .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u');
    }

    // 1) NC usada por OP en BD_Ingresos, en TODOS los libros (2024/2025/2026).
    var porPac = {};   // keyNom → { paciente, generado, ncUsada, consumidoLedger, ops:[] }
    function _pac(nom) {
      var k = _cobKeyNom(nom);
      if (!porPac[k]) porPac[k] = { paciente: String(nom || '').trim(), generado: 0, ncUsada: 0, consumidoLedger: 0, ops: [] };
      return porPac[k];
    }

    var anios = [];
    try { for (var a in INGRESOS_IDS) if (INGRESOS_IDS.hasOwnProperty(a)) anios.push(a); } catch (eA) { anios = []; }
    if (!anios.length) anios = [String(new Date().getFullYear())];

    var libros = {};   // ssId → [años] (varios años pueden compartir libro)
    anios.forEach(function (an) {
      var id = (typeof _ingIdDeAnio === 'function') ? _ingIdDeAnio(an) : INGRESOS_IDS[an];
      if (!id) return;
      if (!libros[id]) libros[id] = [];
      libros[id].push(an);
    });

    Object.keys(libros).forEach(function (ssId) {
      var sh;
      try { sh = SpreadsheetApp.openById(ssId).getSheetByName(BD_INGRESOS_TAB); } catch (eO) { return; }
      if (!sh || sh.getLastRow() < 2) return;
      var lastCol = sh.getLastColumn();
      var raw = sh.getDataRange().getValues();
      var hdrs = raw[0].map(function (h) { return String(h).trim().toLowerCase(); });
      var iPD = hdrs.indexOf('pagosdetalle');   // desglose de pago mixto (JSON)

      for (var r = 1; r < raw.length; r++) {
        var op = String(raw[r][0] || '').trim(); if (!op) continue;
        var linea = _cobNum(raw[r][1]);
        var nom = raw[r][3];
        var fp = String(raw[r][12] || '');
        var pagadoLinea = _cobNum(raw[r][10]);
        var ncOp = 0;

        // Pago MIXTO: el desglose vive solo en la línea 1 de la OP (JSON PagosDetalle).
        var pd = (iPD > -1 && linea === 1) ? String(raw[r][iPD] || '').trim() : '';
        if (pd) {
          try {
            var arr = JSON.parse(pd);
            for (var p = 0; p < arr.length; p++) {
              if (_normFP(arr[p].fp) === FP_NC) ncOp += _cobNum(arr[p].monto);
            }
          } catch (eJ) { /* JSON corrupto → se ignora, se reporta abajo por FormaPago */ }
        } else if (_normFP(fp) === FP_NC) {
          // Pago simple 100% NC: cada línea aporta su Pagado.
          ncOp = pagadoLinea;
        }
        if (ncOp <= 0.01) continue;
        var reg = _pac(nom);
        reg.ncUsada += ncOp;
        reg.ops.push({ op: op, monto: Math.round(ncOp * 100) / 100, libro: libros[ssId].join('/') });
      }
    });

    // 2) Crédito GENERADO por paciente (Creditos_Favor: MontoCredito ya es el saldo
    //    vivo; los renglones ':nc-rev' son reversos de NC, se marcan aparte).
    var shCred = SpreadsheetApp.openById(INGRESOS_SS_ID).getSheetByName(COBRANZA_CREDITOS);
    var vivoPorPac = {}, revPorPac = {};
    if (shCred) {
      var cr = shCred.getDataRange().getValues();
      for (var i = 1; i < cr.length; i++) {
        var m = _cobNum(cr[i][3]); if (m <= 0.01) continue;
        var k = _cobKeyNom(cr[i][2]);
        var esRev = String(cr[i][1] || '').indexOf(':nc-rev') > -1;
        if (esRev) revPorPac[k] = (revPorPac[k] || 0) + m;
        else vivoPorPac[k] = (vivoPorPac[k] || 0) + m;
        _pac(cr[i][2]).generado += m;
      }
    }

    // 3) Consumo YA registrado por el ledger nuevo (OPs guardadas después del fix).
    var shCons = SpreadsheetApp.openById(INGRESOS_SS_ID).getSheetByName(COBRANZA_CONSUMO_NC);
    if (shCons) {
      var co = shCons.getDataRange().getValues();
      for (var j = 1; j < co.length; j++) {
        var kc = _cobKeyNom(co[j][2]); if (!kc) continue;
        _pac(co[j][2]).consumidoLedger += _cobNum(co[j][4]) + _cobNum(co[j][5]);
      }
    }

    // 4) Reporte.
    var out = { ok: true, generadoAt: new Date(), pacientes: [], totales: { ncUsada: 0, creditoVivo: 0, noDescontado: 0, sospechosos: 0 } };
    Object.keys(porPac).forEach(function (k) {
      var p = porPac[k];
      if (p.ncUsada <= 0.01) return;                 // solo interesa quien usó NC
      var vivo = vivoPorPac[k] || 0;
      // NC usada que el ledger NUEVO no registró = NC histórica nunca descontada.
      var noDescontado = Math.max(0, p.ncUsada - p.consumidoLedger);
      // ¿Se gastó más NC de la que jamás tuvo de crédito?
      var excede = p.ncUsada > (p.generado + 0.01);
      var rec = {
        paciente: p.paciente,
        creditoGenerado: Math.round(p.generado * 100) / 100,
        creditoVivoHoy: Math.round(vivo * 100) / 100,
        ncUsada: Math.round(p.ncUsada * 100) / 100,
        ncYaDescontada: Math.round(p.consumidoLedger * 100) / 100,
        ncNoDescontada: Math.round(noDescontado * 100) / 100,
        excedeCredito: excede,
        faltante: excede ? Math.round((p.ncUsada - p.generado) * 100) / 100 : 0,
        ops: p.ops
      };
      out.pacientes.push(rec);
      out.totales.ncUsada += rec.ncUsada;
      out.totales.creditoVivo += rec.creditoVivoHoy;
      out.totales.noDescontado += rec.ncNoDescontada;
      if (excede) out.totales.sospechosos++;
    });
    out.pacientes.sort(function (a, b) { return b.ncUsada - a.ncUsada; });
    ['ncUsada', 'creditoVivo', 'noDescontado'].forEach(function (t) { out.totales[t] = Math.round(out.totales[t] * 100) / 100; });

    function _m(n) { return '$' + (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
    var L = [];
    L.push('══════════ AUDITORÍA DE NOTAS DE CRÉDITO ══════════');
    L.push('Generado: ' + out.generadoAt);
    L.push('SOLO LECTURA — no se corrigió ningún dato.');
    L.push('');
    L.push('Pacientes que pagaron con Nota de Crédito: ' + out.pacientes.length);
    L.push('NC usada (total):            ' + _m(out.totales.ncUsada));
    L.push('NC nunca descontada:         ' + _m(out.totales.noDescontado) + '   ← el bug histórico');
    L.push('Crédito a favor vivo hoy:    ' + _m(out.totales.creditoVivo));
    L.push('Pacientes con NC > crédito:  ' + out.totales.sospechosos + (out.totales.sospechosos ? '   ⚠️ REVISAR' : ''));
    L.push('');
    if (!out.pacientes.length) L.push('No hay ninguna operación pagada con Nota de Crédito. Nada que revisar.');
    out.pacientes.forEach(function (p) {
      L.push('──────────────────────────────────────────');
      L.push((p.excedeCredito ? '⚠️  ' : '    ') + p.paciente);
      L.push('     crédito generado ' + _m(p.creditoGenerado) + ' · vivo hoy ' + _m(p.creditoVivoHoy));
      L.push('     NC usada ' + _m(p.ncUsada) + '  (ya descontada ' + _m(p.ncYaDescontada) + ' · SIN descontar ' + _m(p.ncNoDescontada) + ')');
      if (p.excedeCredito) L.push('     ⚠️ GASTÓ ' + _m(p.faltante) + ' DE CRÉDITO QUE NUNCA TUVO');
      L.push('     OPs: ' + p.ops.map(function (o) { return o.op + ' ' + _m(o.monto); }).join(' · '));
    });
    L.push('══════════════════════════════════════════');
    Logger.log(L.join('\n'));
    out.reporte = L.join('\n');
    return out;
  } catch (ex) {
    Logger.log('auditarNotasCredito ERROR: ' + ex.message);
    return { ok: false, error: ex.message };
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

/* ═════════════════ CONCILIACIÓN MERCADO PAGO (Suscriptores) ═════════════════
 * Cruza la exportación de "Suscriptores" de Mercado Pago contra el universo de
 * suscripciones de crío del ERP (Motor B). El objetivo: saber quién está de verdad
 * dado de alta en la suscripción recurrente de MP (cobro automático) vs. quién solo
 * pagó a mano (transferencia/efectivo) y NO se le va a cobrar solo.
 *
 * VÍNCULO NO DESTRUCTIVO: la liga payer_id ↔ paciente se guarda en una hoja NUEVA
 * y aparte, `Suscripciones_MP` (nunca toca ni Cuentas_Cobrar ni Inventario Crío).
 * Las columnas se agregan al FINAL (patrón _ingColEnsure): jamás se reordenan ni se
 * quitan columnas existentes. Re-subir el archivo es idempotente: primero empata por
 * payer_id (si ya está ligado → solo refresca estatus), si no por nombre; ACTUALIZA la
 * misma fila (nunca duplica). Es solo sincronización: no rompe la base de cobranza.
 */
var COBRANZA_MP        = 'Suscripciones_MP';
var COBRANZA_MP_VER    = 'crio-mp-2026.07.14';
var COBRANZA_MP_HEADERS = ['PacienteKey','Paciente','PayerID','MP_Email','MP_Status','MP_Plan','MP_Monto',
  'MP_Inicio','MP_ProxCobro','MP_UltCobro','MP_UltMonto','MP_Cobros','Clasificacion','ActualizadoEn','ActualizadoPor'];

function _mpEnsureSheet() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COBRANZA_MP);
  if (!sh) {
    sh = ss.insertSheet(COBRANZA_MP);
    sh.getRange(1, 1, 1, COBRANZA_MP_HEADERS.length).setValues([COBRANZA_MP_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
    sh.setFrozenRows(1);
  }
  return sh;
}
// Índice (1-based) de la columna `want` por coincidencia de encabezado. Si no existe,
// la AGREGA al final sin tocar las columnas previas (mismo patrón que _ingColEnsure).
function _mpColEnsure(sh, want, headerText) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
  want = String(want).toLowerCase();
  for (var c = 0; c < hdrs.length; c++) { if (hdrs[c] === want) return c + 1; }
  sh.getRange(1, lastCol + 1).setValue(headerText || want);
  sh.getRange(1, lastCol + 1).setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  return lastCol + 1;
}
// Mapa {clave lógica → índice 1-based}. Asegura (append-safe) todas las columnas que
// usamos, de modo que una hoja creada por una versión previa se migra sin perder datos.
function _mpColMap(sh) {
  return {
    key:    _mpColEnsure(sh, 'pacientekey', 'PacienteKey'),
    pac:    _mpColEnsure(sh, 'paciente', 'Paciente'),
    payer:  _mpColEnsure(sh, 'payerid', 'PayerID'),
    email:  _mpColEnsure(sh, 'mp_email', 'MP_Email'),
    status: _mpColEnsure(sh, 'mp_status', 'MP_Status'),
    plan:   _mpColEnsure(sh, 'mp_plan', 'MP_Plan'),
    monto:  _mpColEnsure(sh, 'mp_monto', 'MP_Monto'),
    inicio: _mpColEnsure(sh, 'mp_inicio', 'MP_Inicio'),
    prox:   _mpColEnsure(sh, 'mp_proxcobro', 'MP_ProxCobro'),
    ultC:   _mpColEnsure(sh, 'mp_ultcobro', 'MP_UltCobro'),
    ultM:   _mpColEnsure(sh, 'mp_ultmonto', 'MP_UltMonto'),
    cobros: _mpColEnsure(sh, 'mp_cobros', 'MP_Cobros'),
    clase:  _mpColEnsure(sh, 'clasificacion', 'Clasificacion'),
    upd:    _mpColEnsure(sh, 'actualizadoen', 'ActualizadoEn'),
    updBy:  _mpColEnsure(sh, 'actualizadopor', 'ActualizadoPor')
  };
}
// Lee todas las ligas persistidas → índices por payer_id y por paciente (idempotencia).
function _mpReadLinks() {
  var sh = _mpEnsureSheet();
  var cm = _mpColMap(sh);
  var out = { sh: sh, cm: cm, byPayer: {}, byKey: {}, rows: [] };
  var lr = sh.getLastRow();
  if (lr < 2) return out;
  var lc = sh.getLastColumn();
  var raw = sh.getRange(1, 1, lr, lc).getValues();
  for (var i = 1; i < raw.length; i++) {
    var r = raw[i];
    function cell(idx) { return (idx && idx <= r.length) ? r[idx - 1] : ''; }
    var payer = String(cell(cm.payer) || '').trim();
    var key = String(cell(cm.key) || '').trim();
    if (!payer && !key) continue;
    var rec = {
      rowNum: i + 1, pacKey: key, paciente: String(cell(cm.pac) || ''), payerId: payer,
      mpEmail: String(cell(cm.email) || ''), mpStatus: String(cell(cm.status) || ''),
      mpPlan: String(cell(cm.plan) || ''), mpMonto: _cobNum(cell(cm.monto)),
      mpInicio: _cobStr(_cobD(cell(cm.inicio))), mpProx: _cobStr(_cobD(cell(cm.prox))),
      mpUltCobro: _cobStr(_cobD(cell(cm.ultC))), mpUltMonto: _cobNum(cell(cm.ultM)),
      mpCobros: _cobNum(cell(cm.cobros)), clasificacion: String(cell(cm.clase) || '')
    };
    out.rows.push(rec);
    if (payer) out.byPayer[payer.toLowerCase()] = rec;
    if (key) out.byKey[key] = rec;
  }
  return out;
}
// Upsert idempotente de una liga. Empata por payer_id, si no por PacienteKey; si no,
// agrega una fila nueva. NUNCA duplica. Actualiza también los índices en memoria.
function _mpUpsertLink(links, rec, usuario) {
  var sh = links.sh, cm = links.cm;
  var existing = null;
  if (rec.payerId && links.byPayer[String(rec.payerId).toLowerCase()]) existing = links.byPayer[String(rec.payerId).toLowerCase()];
  else if (rec.pacKey && links.byKey[rec.pacKey]) existing = links.byKey[rec.pacKey];
  var lc = Math.max(sh.getLastColumn(), COBRANZA_MP_HEADERS.length);
  var rowNum, arr;
  if (existing) {
    rowNum = existing.rowNum;
    arr = sh.getRange(rowNum, 1, 1, lc).getValues()[0];
  } else {
    rowNum = sh.getLastRow() + 1;
    arr = []; for (var z = 0; z < lc; z++) arr.push('');
  }
  function setC(idx, val) { if (idx && idx <= arr.length && val !== undefined && val !== null) arr[idx - 1] = val; }
  // Si ya existía y no traemos payer/campo, respetamos lo que ya había (no lo borramos).
  if (rec.pacKey) setC(cm.key, rec.pacKey);
  if (rec.paciente) setC(cm.pac, rec.paciente);
  if (rec.payerId) setC(cm.payer, rec.payerId);
  if (rec.mpEmail !== undefined && rec.mpEmail !== '') setC(cm.email, rec.mpEmail);
  if (rec.mpStatus !== undefined) setC(cm.status, rec.mpStatus);
  if (rec.mpPlan !== undefined) setC(cm.plan, rec.mpPlan);
  if (rec.mpMonto !== undefined && rec.mpMonto !== '') setC(cm.monto, rec.mpMonto);
  if (rec.mpInicio !== undefined && rec.mpInicio !== '') setC(cm.inicio, rec.mpInicio);
  if (rec.mpProx !== undefined && rec.mpProx !== '') setC(cm.prox, rec.mpProx);
  if (rec.mpUltCobro !== undefined && rec.mpUltCobro !== '') setC(cm.ultC, rec.mpUltCobro);
  if (rec.mpUltMonto !== undefined && rec.mpUltMonto !== '') setC(cm.ultM, rec.mpUltMonto);
  if (rec.mpCobros !== undefined && rec.mpCobros !== '') setC(cm.cobros, rec.mpCobros);
  if (rec.clasificacion !== undefined) setC(cm.clase, rec.clasificacion);
  setC(cm.upd, _cobStr(_cobToday()));
  setC(cm.updBy, usuario || '');
  sh.getRange(rowNum, 1, 1, arr.length).setValues([arr]);
  // refresca índices en memoria para que un mismo run no duplique
  var mem = {
    rowNum: rowNum, pacKey: String(arr[cm.key - 1] || ''), paciente: String(arr[cm.pac - 1] || ''),
    payerId: String(arr[cm.payer - 1] || ''), mpEmail: String(arr[cm.email - 1] || ''),
    mpStatus: String(arr[cm.status - 1] || ''), mpPlan: String(arr[cm.plan - 1] || ''),
    mpMonto: _cobNum(arr[cm.monto - 1]), mpInicio: _cobStr(_cobD(arr[cm.inicio - 1])),
    mpProx: _cobStr(_cobD(arr[cm.prox - 1])), mpUltCobro: _cobStr(_cobD(arr[cm.ultC - 1])),
    mpUltMonto: _cobNum(arr[cm.ultM - 1]), mpCobros: _cobNum(arr[cm.cobros - 1]),
    clasificacion: String(arr[cm.clase - 1] || '')
  };
  if (mem.payerId) links.byPayer[mem.payerId.toLowerCase()] = mem;
  if (mem.pacKey) links.byKey[mem.pacKey] = mem;
  var idx = -1; for (var q = 0; q < links.rows.length; q++) { if (links.rows[q].rowNum === rowNum) { idx = q; break; } }
  if (idx >= 0) links.rows[idx] = mem; else links.rows.push(mem);
  return { rowNum: rowNum, created: !existing };
}

/* ── Normalización de un suscriptor de MP (acepta llaves snake_case o variantes) ── */
function _mpg(obj) {
  if (!obj) return '';
  // índice normalizado (minúsculas, sin no-alfanuméricos) una sola vez
  if (!obj.__mpidx) {
    var idx = {}; for (var k in obj) { if (k === '__mpidx') continue; idx[String(k).toLowerCase().replace(/[^a-z0-9]/g, '')] = obj[k]; }
    obj.__mpidx = idx;
  }
  for (var a = 1; a < arguments.length; a++) {
    var v = obj.__mpidx[arguments[a]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return v;
  }
  return '';
}
function _mpParseSub(raw) {
  var first = String(_mpg(raw, 'payerfirstname', 'firstname', 'nombre') || '').trim();
  var last = String(_mpg(raw, 'payerlastname', 'lastname', 'apellido') || '').trim();
  var nombreMP = (first + ' ' + last).trim();
  return {
    id: String(_mpg(raw, 'id') || '').trim(),
    payerId: String(_mpg(raw, 'payerid') || '').trim(),
    first: first, last: last, nombreMP: nombreMP,
    status: String(_mpg(raw, 'status', 'estado') || '').trim().toLowerCase(),
    reason: String(_mpg(raw, 'reason', 'motivo', 'razon') || '').trim(),
    externalRef: String(_mpg(raw, 'externalreference', 'referencia') || '').trim(),
    planId: String(_mpg(raw, 'preapprovalplanid', 'planid') || '').trim(),
    frequency: _cobNum(_mpg(raw, 'frequency')),
    frequencyType: String(_mpg(raw, 'frequencytype') || '').trim().toLowerCase(),
    monto: _cobNum(_mpg(raw, 'transactionamount', 'amount', 'monto', 'chargeamount')),
    startDate: _cobStr(_cobD(_mpg(raw, 'startdate'))),
    nextPayment: _cobStr(_cobD(_mpg(raw, 'nextpaymentdate'))),
    lastChargeDate: _cobStr(_cobD(_mpg(raw, 'lastchargedate'))),
    lastChargeAmount: _cobNum(_mpg(raw, 'lastchargeamount')),
    chargedQty: _cobNum(_mpg(raw, 'chargedquantity')),
    pendingQty: _cobNum(_mpg(raw, 'pendingchargequantity')),
    email: String(_mpg(raw, 'payeremail', 'email', 'correo') || '').trim(),
    billingStatus: String(_mpg(raw, 'billingstatus') || '').trim().toLowerCase()
  };
}
// Plan MP: anual (~tarifaAnual) vs mensual (~tarifaMensual), por frecuencia o monto.
function _mpPlan(sub, cfg) {
  if (sub.frequencyType && (sub.frequencyType.indexOf('month') > -1 || sub.frequencyType.indexOf('mes') > -1)) {
    if (sub.frequency >= 12) return 'anual';
    if (sub.frequency === 1) return 'mensual';
  }
  if (sub.frequencyType && (sub.frequencyType.indexOf('year') > -1 || sub.frequencyType.indexOf('año') > -1 || sub.frequencyType.indexOf('anio') > -1)) return 'anual';
  var m = sub.monto || 0;
  if (m > 0) {
    var dA = Math.abs(m - (cfg.tarifaAnual || 5700)), dM = Math.abs(m - (cfg.tarifaMensual || 475));
    return dA <= dM ? 'anual' : 'mensual';
  }
  return 'anual';
}
// Estado del suscriptor visto desde MP.
function _mpClaseMP(sub, today) {
  var authorized = (sub.status === 'authorized');
  var hasCharged = (sub.lastChargeAmount > 0) || (sub.chargedQty > 0) || !!sub.lastChargeDate;
  var firstCharge = sub.nextPayment || sub.startDate;
  var fcD = _cobD(firstCharge);
  var future = fcD && fcD.getTime() > today.getTime();
  if (!authorized) return 'inactiva';                 // pausada / cancelada / pendiente
  if (hasCharged) return 'cobrando';                  // ya cobra automáticamente
  if (future) return 'anio-gratis';                   // autorizada, primer cobro a futuro (año gratis)
  return 'por-cobrar';                                // autorizada, sin cobro aún, primer cobro inminente
}
function _mpFirstCharge(sub) { return sub.nextPayment || sub.startDate || ''; }

// Convierte una liga persistida (fila) a una forma "sub" para clasificar sin re-subir.
function _mpLinkToSub(link, today) {
  var sub = {
    payerId: link.payerId, nombreMP: link.paciente, status: String(link.mpStatus || '').toLowerCase(),
    plan: link.mpPlan, monto: link.mpMonto, startDate: link.mpInicio, nextPayment: link.mpProx,
    lastChargeDate: link.mpUltCobro, lastChargeAmount: link.mpUltMonto, chargedQty: link.mpCobros,
    email: link.mpEmail
  };
  sub.claseMP = link.clasificacion || _mpClaseMP(sub, today);
  return sub;
}

/* ── Motor de conciliación (upload) o de estado (solo lectura de ligas) ── */
function conciliarSuscripcionesMP(body) {
  try {
    body = body || {};
    var soloEstado = !!body.soloEstado;
    var canWrite = !(typeof _tokenHasPermission === 'function') || _tokenHasPermission(body.token || '', 'editar_egresos');
    var subsRaw = body.suscriptores || body.subs || [];
    if (!soloEstado && (!subsRaw || !subsRaw.length))
      return { ok: false, error: 'El archivo no traía suscriptores legibles.', version: COBRANZA_MP_VER };

    var cfg = _cobCfg(), today = _cobToday();
    // Universo ERP de crío con su cálculo (Motor B), por llave de paciente.
    var crioData = _cobReadCrio();
    if (crioData.error) return { ok: false, error: crioData.error, version: COBRANZA_MP_VER };
    var pagos = _cobStoragePagosByPac();
    var abonos = _cobReadAbonos();
    var erpByKey = {}, erpList = [];
    for (var i = 0; i < crioData.pacientes.length; i++) {
      var pac = crioData.pacientes[i];
      var calc = _cobCalcSuscripcion(pac, pagos[pac.key] || [], abonos.susByPac[pac.key] || [], cfg, today);
      var rec = {
        key: pac.key, nombre: pac.nombre, estatus: calc.estatus, plan: calc.plan,
        proximoCobro: calc.proximoCobro, billStart: calc.billStart, ultimoPago: calc.ultimoPago,
        montoDebe: calc.montoDebe, esExterno: calc.esExterno, tienePago: !!calc.ultimoPago,
        oov: pac.oov, emb: pac.emb
      };
      erpByKey[pac.key] = rec; erpList.push(rec);
    }

    var links = _mpReadLinks();
    var mpByKey = {};   // pacKey → sub (de esta subida)
    var sinMatch = [];  // suscriptores MP sin paciente en el ERP
    var updated = 0;

    // 1) Procesar suscriptores subidos: clasificar, empatar, y persistir la liga.
    for (var s = 0; s < (subsRaw ? subsRaw.length : 0); s++) {
      var sub = _mpParseSub(subsRaw[s]);
      if (!sub.payerId && !sub.nombreMP) continue;
      sub.claseMP = _mpClaseMP(sub, today);
      sub.plan = _mpPlan(sub, cfg);
      // (a) por payer_id ya ligado, (b) por nombre exacto, (c) difuso.
      var matchedKey = '';
      var lk = sub.payerId ? links.byPayer[sub.payerId.toLowerCase()] : null;
      if (lk && lk.pacKey && erpByKey[lk.pacKey]) matchedKey = lk.pacKey;
      if (!matchedKey) {
        var nk = _cobKeyNom(sub.nombreMP);
        if (nk && erpByKey[nk]) matchedKey = nk;
        else matchedKey = _mpFuzzyMatch(sub, erpByKey);
      }
      if (matchedKey) {
        mpByKey[matchedKey] = sub;
        if (canWrite) {
          _mpUpsertLink(links, {
            pacKey: matchedKey, paciente: erpByKey[matchedKey].nombre, payerId: sub.payerId,
            mpEmail: sub.email, mpStatus: sub.status, mpPlan: sub.plan, mpMonto: sub.monto,
            mpInicio: sub.startDate, mpProx: sub.nextPayment, mpUltCobro: sub.lastChargeDate,
            mpUltMonto: sub.lastChargeAmount, mpCobros: sub.chargedQty, clasificacion: sub.claseMP
          }, body.usuario || '');
          updated++;
        }
      } else {
        sinMatch.push({
          payerId: sub.payerId, nombre: sub.nombreMP, status: sub.status, clase: sub.claseMP,
          plan: sub.plan, monto: sub.monto, inicio: sub.startDate, proxCobro: sub.nextPayment,
          ultCobro: sub.lastChargeDate, ref: sub.externalRef
        });
      }
    }

    // 2) Para cada paciente de crío del ERP, determinar su MP (subida o liga persistida)
    //    y clasificar en las variantes.
    function mpForKey(key) {
      if (mpByKey[key]) return mpByKey[key];
      var l = links.byKey[key];
      if (l && (l.payerId || l.mpStatus)) return _mpLinkToSub(l, today);
      return null;
    }
    var buckets = { cobrando: 0, mpAnioGratis: 0, pagoSinMP: 0, anioGratisSinMP: 0, faltaSinMP: 0, sinMatchERP: sinMatch.length };
    var pacientes = [];
    for (var e = 0; e < erpList.length; e++) {
      var er = erpList[e];
      var mp = mpForKey(er.key);
      var enAnioGratis = (er.estatus === 'Cortesía');
      var bAnio = _cobD(er.billStart);
      if (!enAnioGratis && bAnio && bAnio.getTime() > today.getTime() && !er.tienePago) enAnioGratis = true;
      var bucket, accion = '', accionFecha = '';
      var mpActivo = mp && (mp.status === 'authorized');
      if (mpActivo && mp.claseMP === 'cobrando') {
        bucket = 'cobrando';
      } else if (mpActivo && (mp.claseMP === 'anio-gratis' || mp.claseMP === 'por-cobrar')) {
        bucket = 'mpAnioGratis';
        accionFecha = _mpFirstCharge(mp);
        accion = 'Sin acción — MP cobra solo el ' + (accionFecha || '¿fecha?');
      } else if (enAnioGratis) {
        bucket = 'anioGratisSinMP';
        accionFecha = er.proximoCobro || er.billStart || '';
        accion = 'Registrar en MP o recordar cobro al terminar año gratis (' + (accionFecha || '¿fecha?') + ')';
      } else if (er.tienePago) {
        bucket = 'pagoSinMP';
        accionFecha = er.proximoCobro || '';
        accion = 'No se cobra solo — registrar en MP o recordar cobro (' + (accionFecha || '¿fecha?') + ')';
      } else {
        bucket = 'faltaSinMP';
        accionFecha = er.proximoCobro || '';
        accion = 'Sin suscripción y sin pago — dar de alta en MP o cobrar (' + (accionFecha || '¿fecha?') + ')';
      }
      buckets[bucket] = (buckets[bucket] || 0) + 1;
      pacientes.push({
        nombre: _cobMasked() ? ('Paciente ' + (e + 1)) : er.nombre, key: _cobMasked() ? '' : er.key,
        bucket: bucket, estatusERP: er.estatus, plan: (mp && mp.plan) ? mp.plan : er.plan,
        payerId: mp ? (mp.payerId || '') : '', mpEmail: mp ? (mp.email || '') : '',
        mpStatus: mp ? (mp.status || '') : '', mpClase: mp ? (mp.claseMP || '') : '',
        proximoCobroERP: er.proximoCobro, primerCobroMP: mp ? _mpFirstCharge(mp) : '',
        montoMP: mp ? (mp.monto || 0) : 0, ultimoCobroMP: mp ? (mp.lastChargeDate || '') : '',
        ultimoPagoERP: er.ultimoPago, montoDebe: er.montoDebe, tipoPaciente: er.esExterno ? 'Externo' : 'Hestia',
        oov: er.oov, emb: er.emb, accion: accion, accionFecha: accionFecha
      });
    }
    // orden: primero los que requieren acción
    var rank = { pagoSinMP: 0, anioGratisSinMP: 1, faltaSinMP: 2, mpAnioGratis: 3, cobrando: 4 };
    pacientes.sort(function (a, b) {
      var ra = rank[a.bucket] === undefined ? 9 : rank[a.bucket];
      var rb = rank[b.bucket] === undefined ? 9 : rank[b.bucket];
      if (ra !== rb) return ra - rb;
      return b.montoDebe - a.montoDebe;
    });

    return {
      ok: true, version: COBRANZA_MP_VER, view: 'conciliar-crio-mp',
      resumen: buckets, total: pacientes.length, pacientes: pacientes, sinMatch: sinMatch,
      updated: updated, canWrite: canWrite, masked: _cobMasked(),
      tarifaAnual: cfg.tarifaAnual, tarifaMensual: cfg.tarifaMensual
    };
  } catch (ex) {
    return { ok: false, error: ex.message, version: COBRANZA_MP_VER };
  }
}
// Empate difuso: mismo nombre invertido (apellido+nombre) o contención de nombres.
function _mpFuzzyMatch(sub, erpByKey) {
  var a = _cobKeyNom((sub.last + ' ' + sub.first).trim());
  if (a && erpByKey[a]) return a;
  var ref = _cobKeyNom(sub.externalRef);
  if (ref && erpByKey[ref]) return ref;
  var target = _cobKeyNom(sub.nombreMP);
  if (!target || target.length < 5) return '';
  var tw = target.split(' ').filter(function (w) { return w.length >= 3; });
  if (tw.length < 2) return '';
  var best = '';
  for (var k in erpByKey) {
    var kw = k.split(' ');
    var hits = 0;
    for (var i = 0; i < tw.length; i++) { if (kw.indexOf(tw[i]) > -1) hits++; }
    if (hits >= 2 && hits === tw.length) { if (best) return ''; best = k; } // único candidato exacto en palabras
  }
  return best;
}
// Guardar liga manual (payer_id ↔ paciente) confirmada por el usuario.
function guardarLinkMP(body) {
  try {
    body = body || {};
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_egresos'))
      return { ok: false, error: 'Sin permiso (requiere editar cuentas por pagar).', version: COBRANZA_MP_VER };
    var paciente = String(body.paciente || '').trim();
    var payerId = String(body.payer_id || body.payerId || '').trim();
    if (!paciente) return { ok: false, error: 'Falta el paciente a vincular.', version: COBRANZA_MP_VER };
    if (!payerId) return { ok: false, error: 'Falta el payer_id de MP.', version: COBRANZA_MP_VER };
    var links = _mpReadLinks();
    var up = _mpUpsertLink(links, {
      pacKey: _cobKeyNom(paciente), paciente: paciente, payerId: payerId,
      mpEmail: body.mp_email || '', mpStatus: body.mp_status || '', mpPlan: body.mp_plan || '',
      mpMonto: (body.mp_monto === undefined ? '' : body.mp_monto), mpInicio: body.mp_inicio || '',
      mpProx: body.mp_prox || '', mpUltCobro: body.mp_ultcobro || '',
      mpUltMonto: (body.mp_ultmonto === undefined ? '' : body.mp_ultmonto),
      mpCobros: (body.mp_cobros === undefined ? '' : body.mp_cobros), clasificacion: body.clase || ''
    }, body.usuario || '');
    return { ok: true, version: COBRANZA_MP_VER, rowNum: up.rowNum, created: up.created, paciente: paciente, payerId: payerId };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_MP_VER }; }
}
// Deshacer una liga manual (por payer_id o por paciente). No borra datos de cobranza.
function desvincularLinkMP(body) {
  try {
    body = body || {};
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'editar_egresos'))
      return { ok: false, error: 'Sin permiso (requiere editar cuentas por pagar).', version: COBRANZA_MP_VER };
    var payerId = String(body.payer_id || body.payerId || '').trim();
    var paciente = String(body.paciente || '').trim();
    var links = _mpReadLinks();
    var target = null;
    if (payerId && links.byPayer[payerId.toLowerCase()]) target = links.byPayer[payerId.toLowerCase()];
    else if (paciente && links.byKey[_cobKeyNom(paciente)]) target = links.byKey[_cobKeyNom(paciente)];
    if (!target) return { ok: false, error: 'No se encontró la liga a deshacer.', version: COBRANZA_MP_VER };
    links.sh.deleteRow(target.rowNum);
    return { ok: true, version: COBRANZA_MP_VER, borrada: target.rowNum };
  } catch (ex) { return { ok: false, error: ex.message, version: COBRANZA_MP_VER }; }
}
// Inicializa la hoja de ligas MP (idempotente).
function setupConciliacionMP() { _mpEnsureSheet(); return { ok: true, version: COBRANZA_MP_VER, tab: COBRANZA_MP }; }

/* ════════════════════ SUGERIDOR de traducciones ═════════════════════════
 * ⚠ ESTO YA NO ES LA AUTORIDAD. Solo SUGIERE.
 *
 * La traducción con la que se imprime un documento vive en el CATÁLOGO
 * (BD_Productos, columna Descripcion_<LANG> — ver traducciones.gs). Esta
 * función solo propone un texto para que un humano lo confirme o corrija
 * al capturar el producto o desde la pantalla de bloqueo de impresión.
 *
 *   traducirTextos({ lang:'en'|'fr'|'pt', textos:[...] | '["..",".."]' })
 *     → { ok:true, lang, textos:{ original: sugerencia, ... }, faltan:[...] }
 *
 * - `textos` trae SOLO lo que se pudo traducir de verdad.
 * - `faltan` trae lo que NO se pudo: el llamador debe pedirle al humano que
 *   lo escriba. NUNCA se devuelve el español como si fuera traducción — ese
 *   passthrough mudo era la causa de las cartas mitad inglés / mitad español.
 * - SEG_TRAD_OVERRIDES: términos de fertilidad curados a mano. Hoy sirven de
 *   SEMILLA del catálogo (migrarTraduccionesProductos) y de sugerencia; ya no
 *   deciden lo que se imprime.
 * - Cache en CacheService (6 h) por texto normalizado + idioma.
 */
var SEG_TRAD_OVERRIDES = {
  en: {
    'fertilizacion in vitro':'In Vitro Fertilization (IVF)',
    'fertilizacion in vitro (fiv)':'In Vitro Fertilization (IVF)',
    'fiv':'In Vitro Fertilization (IVF)',
    'histeroscopia':'Hysteroscopy',
    'histeroscopia diagnostica':'Diagnostic Hysteroscopy',
    'laparoscopia':'Laparoscopy',
    'criopreservacion':'Cryopreservation',
    'criopreservacion de embriones':'Embryo Cryopreservation',
    'criopreservacion de ovulos':'Egg Cryopreservation',
    'vitrificacion de ovulos':'Egg Vitrification',
    'estimulacion ovarica controlada':'Controlled Ovarian Stimulation',
    'estimulacion ovarica':'Ovarian Stimulation',
    'transferencia de embriones':'Embryo Transfer',
    'transferencia de embrion':'Embryo Transfer',
    'transferencia de embriones congelados':'Frozen Embryo Transfer',
    'inseminacion artificial':'Artificial Insemination',
    'inseminacion intrauterina':'Intrauterine Insemination (IUI)',
    'icsi':'Intracytoplasmic Sperm Injection (ICSI)',
    'aspiracion folicular':'Follicular Aspiration',
    'captura ovocitaria':'Oocyte Retrieval',
    'aspiracion de ovulos':'Oocyte Retrieval',
    'donacion de ovulos':'Egg Donation',
    'ovodonacion':'Egg Donation',
    'donacion de esperma':'Sperm Donation',
    'biopsia embrionaria':'Embryo Biopsy',
    'estudio genetico preimplantacional':'Preimplantation Genetic Testing (PGT)',
    'diagnostico genetico preimplantacional':'Preimplantation Genetic Diagnosis (PGD)',
    'almacenamiento de embriones':'Embryo Storage',
    'almacenamiento crio':'Cryogenic Storage',
    'renta de vientre':'Gestational Surrogacy',
    'gestacion subrogada':'Gestational Surrogacy',
    'consulta':'Consultation',
    'consulta medica':'Medical Consultation',
    'ultrasonido':'Ultrasound',
    'laboratorio':'Laboratory',
    'medicamentos':'Medications'
  },
  fr: {
    'fertilizacion in vitro':'Fécondation in vitro (FIV)',
    'fiv':'Fécondation in vitro (FIV)',
    'histeroscopia':'Hystéroscopie',
    'criopreservacion':'Cryoconservation',
    'criopreservacion de embriones':"Cryoconservation d'embryons",
    'estimulacion ovarica controlada':'Stimulation ovarienne contrôlée',
    'transferencia de embriones':"Transfert d'embryons",
    'inseminacion artificial':'Insémination artificielle',
    'icsi':'Injection intracytoplasmique de spermatozoïdes (ICSI)',
    'donacion de ovulos':"Don d'ovocytes",
    'consulta':'Consultation',
    'ultrasonido':'Échographie',
    'medicamentos':'Médicaments'
  },
  pt: {
    'fertilizacion in vitro':'Fertilização in vitro (FIV)',
    'fiv':'Fertilização in vitro (FIV)',
    'histeroscopia':'Histeroscopia',
    'criopreservacion':'Criopreservação',
    'criopreservacion de embriones':'Criopreservação de embriões',
    'estimulacion ovarica controlada':'Estimulação ovariana controlada',
    'transferencia de embriones':'Transferência de embriões',
    'inseminacion artificial':'Inseminação artificial',
    'icsi':'Injeção intracitoplasmática de espermatozoides (ICSI)',
    'donacion de ovulos':'Doação de óvulos',
    'consulta':'Consulta',
    'ultrasonido':'Ultrassom',
    'medicamentos':'Medicamentos'
  }
};
// Normaliza para comparar contra el mapa curado y para la clave de cache:
// minúsculas, sin acentos, espacios colapsados (sin depender de String.normalize).
function _segTrNorm(s) {
  s = String(s || '').toLowerCase();
  var map = { 'á':'a','é':'e','í':'i','ó':'o','ú':'u','ü':'u','ñ':'n' };
  s = s.replace(/[áéíóúüñ]/g, function (c) { return map[c] || c; });
  return s.replace(/\s+/g, ' ').trim();
}
function _segTrHash(s) {
  var h = 0;
  for (var i = 0; i < s.length; i++) { h = ((h << 5) - h + s.charCodeAt(i)) | 0; }
  return (h >>> 0).toString(36);
}
function traducirTextos(body) {
  try {
    body = body || {};
    var lang = String(body.lang || body.idioma || 'en').toLowerCase().substring(0, 2);
    var textos = body.textos;
    if (typeof textos === 'string') { try { textos = JSON.parse(textos); } catch (e) { textos = textos ? [textos] : []; } }
    if (!textos || !textos.length) return { ok: true, lang: lang, textos: {} };

    var out = {}, faltan = [];
    // Español: el texto original YA está en español, no hay nada que traducir.
    if (lang === 'es') {
      for (var p = 0; p < textos.length; p++) out[textos[p]] = textos[p];
      return { ok: true, lang: lang, textos: out, faltan: [] };
    }
    // Idioma que el traductor no soporta: se dice, NO se devuelve el español
    // haciéndolo pasar por traducción. Un passthrough mudo aquí es exactamente
    // lo que producía cartas en inglés con los servicios en español.
    if (lang !== 'en' && lang !== 'fr' && lang !== 'pt') {
      return { ok: false, lang: lang, error: 'Idioma no soportado por el sugeridor: ' + lang,
               textos: {}, faltan: textos.slice() };
    }

    var over = SEG_TRAD_OVERRIDES[lang] || {};
    var cache = null; try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
    var pend = [], pendKey = [];

    for (var i = 0; i < textos.length; i++) {
      var orig = String(textos[i] == null ? '' : textos[i]).trim();
      if (!orig) { out[textos[i]] = textos[i]; continue; }
      if (out.hasOwnProperty(orig)) continue;
      var norm = _segTrNorm(orig);
      if (over[norm]) { out[orig] = over[norm]; continue; }          // término médico curado
      var ck = 'segtr_' + lang + '_' + _segTrHash(norm);
      var hit = cache ? cache.get(ck) : null;
      if (hit != null) { out[orig] = hit; continue; }                 // cache
      pend.push(orig); pendKey.push(ck);
    }

    // Traduce los pendientes (uno por uno; LanguageApp no admite lote real).
    // Si LanguageApp falla o devuelve vacío, el texto NO se resuelve: se reporta
    // en `faltan`. ANTES se hacía `if (!tr) tr = pend[j]`, que devolvía el
    // español disfrazado de traducción y el documento salía mitad y mitad sin
    // que nadie se enterara. Ese fallback mudo se eliminó a propósito.
    for (var j = 0; j < pend.length; j++) {
      var tr = '';
      try { tr = LanguageApp.translate(pend[j], 'es', lang); } catch (e2) { tr = ''; }
      tr = String(tr || '').trim();
      // Si el traductor devuelve exactamente lo mismo que entró, no tradujo nada.
      if (!tr || _segTrNorm(tr) === _segTrNorm(pend[j])) { faltan.push(pend[j]); continue; }
      out[pend[j]] = tr;
      if (cache) { try { cache.put(pendKey[j], tr, 21600); } catch (e3) {} }
    }
    return { ok: true, lang: lang, textos: out, faltan: faltan };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}
