/* ==============================================================
   cxp_creditos.gs — Créditos de Proveedor + Abonos parciales en CxP
   --------------------------------------------------------------
   CxP y Egresos son el mismo renglón de Egresos2026 (ver finance.gs).
   Este archivo agrega, SIN tocar pagarCxP/deleteCxPRow existentes:

   - Creditos_Proveedor: saldo a favor que queda cuando se cancela una
     orden que ya se había pagado (dinero real que salió sin que llegara
     el producto/servicio).
   - Abonos_CxP: cada aplicación de dinero contra una cuenta por pagar
     (pago normal o crédito aplicado) — el saldo pendiente de una orden
     = Monto − suma de sus abonos activos (no reversados).

   Al cancelar una orden pagada, sus abonos se reversan uno por uno:
   los que vinieron de crédito regresan ese crédito a disponible: los
   que fueron pago real (efectivo/banco) se convierten en un crédito
   nuevo. Caso especial: órdenes pagadas ANTES de que existiera este
   sistema (vía pagarCxP directo, sin abonos) — se detecta por
   Pagado=true sin abonos y se trata como un pago completo a convertir.
   ============================================================== */

var CXP_CRED_TAB = 'Creditos_Proveedor';
var CXP_CRED_HEADERS = ['ID', 'Fecha', 'Proveedor', 'Monto', 'MontoDisponible', 'Origen', 'CxPIdOrigen', 'Usuario', 'Notas'];
var CXP_ABONO_TAB = 'Abonos_CxP';
var CXP_ABONO_HEADERS = ['ID', 'Fecha', 'CxPId', 'Proveedor', 'Concepto', 'Monto', 'Origen', 'CreditoId', 'FormaPago', 'Usuario', 'Notas', 'Reversado'];

function _cxpCredSS() { return SpreadsheetApp.openById(EGRESOS_SS_2026); }

function setupCxPCreditosAbonos() {
  var ss = _cxpCredSS();
  var credSh = ss.getSheetByName(CXP_CRED_TAB);
  if (!credSh) {
    credSh = ss.insertSheet(CXP_CRED_TAB);
    credSh.getRange(1, 1, 1, CXP_CRED_HEADERS.length).setValues([CXP_CRED_HEADERS]);
    credSh.getRange(1, 1, 1, CXP_CRED_HEADERS.length).setFontWeight('bold').setBackground('#166534').setFontColor('#ffffff');
    credSh.setFrozenRows(1);
    credSh.autoResizeColumns(1, CXP_CRED_HEADERS.length);
  }
  var abonoSh = ss.getSheetByName(CXP_ABONO_TAB);
  if (!abonoSh) {
    abonoSh = ss.insertSheet(CXP_ABONO_TAB);
    abonoSh.getRange(1, 1, 1, CXP_ABONO_HEADERS.length).setValues([CXP_ABONO_HEADERS]);
    abonoSh.getRange(1, 1, 1, CXP_ABONO_HEADERS.length).setFontWeight('bold').setBackground('#1a252f').setFontColor('#ffffff');
    abonoSh.setFrozenRows(1);
    abonoSh.autoResizeColumns(1, CXP_ABONO_HEADERS.length);
  }
  return { ok: true, creditosCreada: !ss.getSheetByName(CXP_CRED_TAB) ? false : true, abonosCreada: !ss.getSheetByName(CXP_ABONO_TAB) ? false : true };
}

function _cxpColIdx(headers, name) { var i = headers.indexOf(name); if (i < 0) throw new Error('Columna ' + name + ' no encontrada (¿corriste setupCxPCreditosAbonos()?)'); return i; }

function _cxpNextId(sh, idColName) {
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var idCol = _cxpColIdx(hdrs, idColName);
  var max = 0;
  for (var i = 1; i < data.length; i++) { var n = parseInt(data[i][idCol], 10); if (n > max) max = n; }
  return max + 1;
}

/* ── Créditos de proveedor ─────────────────────────────────────── */
function readCreditosProveedor(proveedor) {
  try {
    var sh = _cxpCredSS().getSheetByName(CXP_CRED_TAB);
    if (!sh) return { ok: true, rows: [] };
    var data = sh.getDataRange().getValues();
    var hdrs = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var disp = Number(r[_cxpColIdx(hdrs, 'MontoDisponible')]) || 0;
      if (disp <= 0.01) continue;
      var prov = String(r[_cxpColIdx(hdrs, 'Proveedor')] || '');
      if (proveedor && prov.trim().toLowerCase() !== String(proveedor).trim().toLowerCase()) continue;
      rows.push({
        _rowNum: i + 1, id: r[_cxpColIdx(hdrs, 'ID')], fecha: String(r[_cxpColIdx(hdrs, 'Fecha')] || ''),
        proveedor: prov, monto: Number(r[_cxpColIdx(hdrs, 'Monto')]) || 0, montoDisponible: disp,
        origen: String(r[_cxpColIdx(hdrs, 'Origen')] || ''), cxpIdOrigen: String(r[_cxpColIdx(hdrs, 'CxPIdOrigen')] || '')
      });
    }
    return { ok: true, rows: rows };
  } catch (ex) { return { ok: false, error: ex.message, rows: [] }; }
}

function _crearCreditoProveedor(p) {
  var sh = _cxpCredSS().getSheetByName(CXP_CRED_TAB);
  if (!sh) throw new Error('Falta correr setupCxPCreditosAbonos()');
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var id = _cxpNextId(sh, 'ID');
  var row = new Array(hdrs.length).fill('');
  row[_cxpColIdx(hdrs, 'ID')] = id;
  row[_cxpColIdx(hdrs, 'Fecha')] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  row[_cxpColIdx(hdrs, 'Proveedor')] = p.proveedor || '';
  row[_cxpColIdx(hdrs, 'Monto')] = p.monto || 0;
  row[_cxpColIdx(hdrs, 'MontoDisponible')] = p.monto || 0;
  row[_cxpColIdx(hdrs, 'Origen')] = p.origen || '';
  row[_cxpColIdx(hdrs, 'CxPIdOrigen')] = p.cxpIdOrigen || '';
  row[_cxpColIdx(hdrs, 'Usuario')] = p.usuario || '';
  row[_cxpColIdx(hdrs, 'Notas')] = p.notas || '';
  sh.appendRow(row);
  return id;
}

function _consumirCredito(creditoId, monto) {
  var sh = _cxpCredSS().getSheetByName(CXP_CRED_TAB);
  if (!sh) return { ok: false, error: 'Falta correr setupCxPCreditosAbonos()' };
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var idCol = _cxpColIdx(hdrs, 'ID'), dispCol = _cxpColIdx(hdrs, 'MontoDisponible');
  var provCol = _cxpColIdx(hdrs, 'Proveedor');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) !== String(creditoId)) continue;
    var disp = Number(data[i][dispCol]) || 0;
    if (monto > disp + 0.01) return { ok: false, error: 'El crédito solo tiene $' + disp.toFixed(2) + ' disponible' };
    sh.getRange(i + 1, dispCol + 1).setValue(Math.max(0, disp - monto));
    return { ok: true, proveedor: String(data[i][provCol] || '') };
  }
  return { ok: false, error: 'Crédito no encontrado' };
}
function _cxpNormProv(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }

/* ── Reparación: reversa abonos de crédito CRUZADOS (crédito de un proveedor
   aplicado por error a una orden de OTRO proveedor) y restaura el crédito.
   Corre UNA vez desde el editor de Apps Script (o vía POST) para corregir
   casos históricos como el crédito de LEPSI PRISMA abonado a CRISBEN. ── */
function repararAbonosCruzados() {
  try {
    var ss = _cxpCredSS();
    var abSh = ss.getSheetByName(CXP_ABONO_TAB); if (!abSh) return { ok: false, error: 'Sin hoja de abonos' };
    var credSh = ss.getSheetByName(CXP_CRED_TAB); if (!credSh) return { ok: false, error: 'Sin hoja de créditos' };
    var ad = abSh.getDataRange().getValues(), ah = ad[0];
    var cd = credSh.getDataRange().getValues(), ch = cd[0];
    var credProv = {}; var idC = _cxpColIdx(ch, 'ID'), pC = _cxpColIdx(ch, 'Proveedor');
    for (var i = 1; i < cd.length; i++) credProv[String(cd[i][idC])] = String(cd[i][pC] || '');
    var iRev = _cxpColIdx(ah, 'Reversado'), iOrig = _cxpColIdx(ah, 'Origen'), iCred = _cxpColIdx(ah, 'CreditoId'),
        iProv = _cxpColIdx(ah, 'Proveedor'), iMonto = _cxpColIdx(ah, 'Monto'), iCxp = _cxpColIdx(ah, 'CxPId');
    var corregidos = [];
    for (var a = 1; a < ad.length; a++) {
      var rev = ad[a][iRev] === true || String(ad[a][iRev]).toUpperCase() === 'TRUE';
      if (rev) continue;
      if (String(ad[a][iOrig] || '') !== 'credito') continue;
      var credId = String(ad[a][iCred] || ''); if (!credId) continue;
      var provCred = credProv[credId] || '', provOrden = String(ad[a][iProv] || '');
      if (!provCred || !provOrden) continue;
      if (_cxpNormProv(provCred) === _cxpNormProv(provOrden)) continue;   // mismo proveedor: ok
      var monto = Number(ad[a][iMonto]) || 0;
      abSh.getRange(a + 1, iRev + 1).setValue(true);       // marca el abono reversado
      _restaurarCredito(credId, monto);                    // devuelve el saldo al crédito original
      corregidos.push({ cxpId: String(ad[a][iCxp] || ''), ordenProveedor: provOrden, creditoDe: provCred, monto: monto });
    }
    try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (e) {}
    try { logAudit('sistema', 'CxP', 'Reparar abonos cruzados', '', '', '', corregidos.length + ' reversados'); } catch (e) {}
    return { ok: true, corregidos: corregidos.length, detalle: corregidos };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

function _restaurarCredito(creditoId, monto) {
  var sh = _cxpCredSS().getSheetByName(CXP_CRED_TAB);
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var idCol = _cxpColIdx(hdrs, 'ID'), dispCol = _cxpColIdx(hdrs, 'MontoDisponible'), montoCol = _cxpColIdx(hdrs, 'Monto');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) !== String(creditoId)) continue;
    var disp = Number(data[i][dispCol]) || 0;
    var tope = Number(data[i][montoCol]) || 0;
    sh.getRange(i + 1, dispCol + 1).setValue(Math.min(tope, disp + monto));
    return;
  }
}

/* ── Abonos ───────────────────────────────────────────────────────── */
function _registrarAbono(p) {
  var sh = _cxpCredSS().getSheetByName(CXP_ABONO_TAB);
  if (!sh) throw new Error('Falta correr setupCxPCreditosAbonos()');
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var id = _cxpNextId(sh, 'ID');
  var row = new Array(hdrs.length).fill('');
  row[_cxpColIdx(hdrs, 'ID')] = id;
  row[_cxpColIdx(hdrs, 'Fecha')] = p.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  row[_cxpColIdx(hdrs, 'CxPId')] = p.cxpId;
  row[_cxpColIdx(hdrs, 'Proveedor')] = p.proveedor || '';
  row[_cxpColIdx(hdrs, 'Concepto')] = p.concepto || '';
  row[_cxpColIdx(hdrs, 'Monto')] = p.monto || 0;
  row[_cxpColIdx(hdrs, 'Origen')] = p.origen || 'pago'; // 'pago' | 'credito'
  row[_cxpColIdx(hdrs, 'CreditoId')] = p.creditoId || '';
  row[_cxpColIdx(hdrs, 'FormaPago')] = p.formaPago || '';
  row[_cxpColIdx(hdrs, 'Usuario')] = p.usuario || '';
  row[_cxpColIdx(hdrs, 'Notas')] = p.notas || '';
  row[_cxpColIdx(hdrs, 'Reversado')] = false;
  sh.appendRow(row);
  return id;
}

function _cxpAbonosPorId(cxpId, soloActivos) {
  var sh = _cxpCredSS().getSheetByName(CXP_ABONO_TAB);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var hdrs = data[0];
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[_cxpColIdx(hdrs, 'CxPId')]) !== String(cxpId)) continue;
    var reversado = r[_cxpColIdx(hdrs, 'Reversado')] === true || String(r[_cxpColIdx(hdrs, 'Reversado')]).toUpperCase() === 'TRUE';
    if (soloActivos && reversado) continue;
    out.push({
      _rowNum: i + 1, id: r[_cxpColIdx(hdrs, 'ID')], monto: Number(r[_cxpColIdx(hdrs, 'Monto')]) || 0,
      origen: String(r[_cxpColIdx(hdrs, 'Origen')] || 'pago'), creditoId: String(r[_cxpColIdx(hdrs, 'CreditoId')] || ''),
      reversado: reversado
    });
  }
  return out;
}

function _cxpSumAbonosActivos(cxpId) {
  return _cxpAbonosPorId(cxpId, true).reduce(function (s, a) { return s + a.monto; }, 0);
}

function _marcarAbonoReversado(abonoRowNum) {
  var sh = _cxpCredSS().getSheetByName(CXP_ABONO_TAB);
  if (!sh) return;
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.getRange(abonoRowNum, _cxpColIdx(hdrs, 'Reversado') + 1).setValue(true);
}

/* Saldo pendiente de una o varias órdenes — usado por readBDCxP para
   mostrar en pantalla solo lo que de verdad se debe. */
function readSaldosCxP(cxpIds) {
  try {
    var sh = _cxpCredSS().getSheetByName(CXP_ABONO_TAB);
    var mapa = {};
    (cxpIds || []).forEach(function (id) { mapa[id] = 0; });
    if (sh) {
      var data = sh.getDataRange().getValues();
      var hdrs = data[0];
      for (var i = 1; i < data.length; i++) {
        var r = data[i];
        var id = String(r[_cxpColIdx(hdrs, 'CxPId')]);
        if (!(id in mapa)) continue;
        var reversado = r[_cxpColIdx(hdrs, 'Reversado')] === true || String(r[_cxpColIdx(hdrs, 'Reversado')]).toUpperCase() === 'TRUE';
        if (reversado) continue;
        mapa[id] += Number(r[_cxpColIdx(hdrs, 'Monto')]) || 0;
      }
    }
    return { ok: true, abonadoPorId: mapa };
  } catch (ex) { return { ok: false, error: ex.message, abonadoPorId: {} }; }
}

/* ── Ruteo a banco/caja chica — copia del bloque ya probado en pagarCxP,
   factorizado para poder reusarlo desde aplicarAbonoCxP sin tocar pagarCxP ── */
function _cxpRutearABanco(formaPago, monto, fechaPago, ref) {
  if (!formaPago || monto <= 0) return;
  var mesStr = fechaPago.substring(0, 7);
  if (formaPago === 'Efectivo') {
    try {
      var ccSh = getCajaChicaSheet();
      var ccData = ccSh.getDataRange().getValues();
      var ccHdr = ccData[0].map(function (h) { return String(h).trim().toUpperCase(); });
      var ciF = ccHdr.indexOf('FECHA'), ciC = ccHdr.indexOf('CONCEPTO'), ciS = ccHdr.indexOf('SALIDA');
      var ccRow = -1;
      for (var ci = 1; ci < ccData.length; ci++) {
        if (!String(ccData[ci][ciF] || '').trim() && !String(ccData[ci][ciC] || '').trim()) { ccRow = ci + 1; break; }
      }
      if (ccRow === -1) ccRow = ccSh.getLastRow() + 1;
      ccSh.getRange(ccRow, ciF + 1).setValue(fechaPago);
      ccSh.getRange(ccRow, ciC + 1).setValue(ref);
      ccSh.getRange(ccRow, ciS + 1).setValue(monto);
      SpreadsheetApp.flush();
    } catch (ccErr) { Logger.log('_cxpRutearABanco Efectivo→CajaChica error: ' + ccErr.message); }
    return;
  }
  var banco = '', bankRow = null;
  if (formaPago === 'Santander' || formaPago === 'Transferencia') {
    banco = 'santander'; bankRow = [fechaPago, 0, monto, 0, ref, '', '', '', ''];
  } else if (formaPago === 'AMEX') {
    banco = 'amex'; bankRow = [fechaPago, monto, 0, ref, '', '', '', '', mesStr];
  } else if (formaPago === 'Mercado Pago' || formaPago === 'TDC' || formaPago === 'TDD') {
    banco = 'mercadopago'; bankRow = [mesStr, fechaPago, 0, 0, 0, -monto, 0, false, ref, 'PAGO'];
  }
  if (banco && bankRow) {
    try { saveBankRow(banco, bankRow); } catch (bErr) { Logger.log('_cxpRutearABanco banco error: ' + bErr.message); }
  }
}

/* ── Devolución de dinero del proveedor: ENTRADA a banco/caja ─────────
   Espejo de _cxpRutearABanco pero como ingreso (depósito/reembolso),
   para que la conciliación cuadre con el movimiento que aparecerá en el
   estado de cuenta. Solo se usa cuando el usuario elige "me devolvieron
   el dinero" al cancelar — el modo "saldo a favor" NO toca bancos. ── */
function _cxpRutearDevolucionABanco(formaPago, monto, fecha, ref) {
  if (!formaPago || monto <= 0) return;
  var mesStr = fecha.substring(0, 7);
  if (formaPago === 'Efectivo') {
    try {
      var ccSh = getCajaChicaSheet();
      var ccData = ccSh.getDataRange().getValues();
      var ccHdr = ccData[0].map(function (h) { return String(h).trim().toUpperCase(); });
      var ciF = ccHdr.indexOf('FECHA'), ciC = ccHdr.indexOf('CONCEPTO'), ciE = ccHdr.indexOf('ENTRADA');
      if (ciE < 0) throw new Error('La hoja de Caja Chica no tiene columna ENTRADA');
      var ccRow = -1;
      for (var ci = 1; ci < ccData.length; ci++) {
        if (!String(ccData[ci][ciF] || '').trim() && !String(ccData[ci][ciC] || '').trim()) { ccRow = ci + 1; break; }
      }
      if (ccRow === -1) ccRow = ccSh.getLastRow() + 1;
      ccSh.getRange(ccRow, ciF + 1).setValue(fecha);
      ccSh.getRange(ccRow, ciC + 1).setValue(ref);
      ccSh.getRange(ccRow, ciE + 1).setValue(monto);
      SpreadsheetApp.flush();
    } catch (ccErr) { Logger.log('_cxpRutearDevolucion Efectivo→CajaChica error: ' + ccErr.message); }
    return;
  }
  var banco = '', bankRow = null;
  if (formaPago === 'Santander' || formaPago === 'Transferencia') {
    // Santander: [Fecha, Deposito, Retiro, Saldo, Descripcion, …]
    banco = 'santander'; bankRow = [fecha, monto, 0, 0, ref, '', '', '', ''];
  } else if (formaPago === 'AMEX') {
    // AMEX: cargo negativo = crédito/reembolso a la tarjeta
    banco = 'amex'; bankRow = [fecha, -monto, 0, ref, '', '', '', '', mesStr];
  } else if (formaPago === 'Mercado Pago' || formaPago === 'TDC' || formaPago === 'TDD') {
    // MP: el pago usa -monto; la devolución entra en positivo
    banco = 'mercadopago'; bankRow = [mesStr, fecha, 0, 0, 0, monto, 0, false, ref, 'DEVOLUCIÓN'];
  }
  if (banco && bankRow) {
    try { saveBankRow(banco, bankRow); } catch (bErr) { Logger.log('_cxpRutearDevolucion banco error: ' + bErr.message); }
  }
}

/* ── Abonar (pago parcial y/o crédito) a una cuenta por pagar ────────
   No reemplaza pagarCxP (que sigue sirviendo para "pagar todo de un
   jalón" con ruteo a banco) — este es el camino nuevo para pagos
   parciales y/o aplicar un crédito de proveedor. Si el saldo llega a
   cero, marca la fila Pagado=TRUE igual que pagarCxP. */
function aplicarAbonoCxP(body) {
  try {
    var ss = _cxpCredSS();
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var sh = ss.getSheetByName(egTab);
    if (!sh) return { ok: false, error: 'Hoja Egresos no encontrada' };
    var rowNum = parseInt(body.rowNum);
    if (!rowNum || rowNum < 2) return { ok: false, error: 'Fila inválida' };
    var rowData = sh.getRange(rowNum, 1, 1, 20).getValues()[0];
    var cxpId = String(rowData[0] || '');
    var proveedor = String(rowData[4] || '');
    var concepto = String(rowData[8] || '');
    var monto = parseFloat(rowData[9]) || 0;

    var saldoActual = monto - _cxpSumAbonosActivos(cxpId);
    var montoCredito = Math.max(0, parseFloat(body.montoCredito) || 0);
    var creditoId = body.creditoId || '';
    var montoPagoCash = Math.max(0, parseFloat(body.montoPagoCash) || 0);
    var formaPago = body.formaPago || '';
    var fechaPago = body.fechaPago || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    if (montoCredito <= 0 && montoPagoCash <= 0) return { ok: false, error: 'Indica un monto de crédito y/o de pago' };
    if (montoCredito + montoPagoCash > saldoActual + 0.01) return { ok: false, error: 'El monto a aplicar ($' + (montoCredito + montoPagoCash).toFixed(2) + ') es mayor al saldo pendiente ($' + saldoActual.toFixed(2) + ')' };
    if (montoPagoCash > 0 && !formaPago) return { ok: false, error: 'Indica la forma de pago para la parte en efectivo/banco' };

    if (montoCredito > 0) {
      if (!creditoId) return { ok: false, error: 'Falta indicar qué crédito aplicar' };
      var okCred = _consumirCredito(creditoId, montoCredito);
      if (!okCred.ok) return okCred;
      // BLINDAJE: un crédito solo se puede aplicar a órdenes del MISMO proveedor.
      if (okCred.proveedor && _cxpNormProv(okCred.proveedor) !== _cxpNormProv(proveedor)) {
        _restaurarCredito(creditoId, montoCredito);   // deshace el consumo
        return { ok: false, error: 'Ese crédito es de "' + okCred.proveedor + '" y no se puede aplicar a una orden de "' + proveedor + '". Cada crédito solo aplica al mismo proveedor.' };
      }
      _registrarAbono({ cxpId: cxpId, proveedor: proveedor, concepto: concepto, monto: montoCredito, origen: 'credito', creditoId: creditoId, usuario: body.usuario || '', fecha: fechaPago });
    }
    if (montoPagoCash > 0) {
      _registrarAbono({ cxpId: cxpId, proveedor: proveedor, concepto: concepto, monto: montoPagoCash, origen: 'pago', formaPago: formaPago, usuario: body.usuario || '', fecha: fechaPago });
      _cxpRutearABanco(formaPago, montoPagoCash, fechaPago, concepto + ' · ' + proveedor + ' [Egreso #' + cxpId + ']');
    }

    var nuevoSaldo = saldoActual - montoCredito - montoPagoCash;
    if (nuevoSaldo <= 0.01) {
      sh.getRange(rowNum, 2).setValue(new Date(fechaPago));
      sh.getRange(rowNum, 14).setValue(true);
      sh.getRange(rowNum, 17).setValue(formaPago || 'Crédito de proveedor');
    }
    try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (e) {}
    logAudit(body.usuario || 'sistema', 'CxP', 'Abono', cxpId, 'Saldo', String(saldoActual.toFixed(2)), String(Math.max(0, nuevoSaldo).toFixed(2)));
    return { ok: true, saldoPendiente: Math.max(0, nuevoSaldo), pagadoCompleto: nuevoSaldo <= 0.01 };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Traza de una orden cancelada: reconstruye a dónde fue su dinero,
   para que el reporte/detalle muestre la historia completa —
   "este pedido nunca se surtió; su $X quedó como crédito, aplicado a la
   orden #Y junto con un depósito de $Z por la diferencia". ──────────── */
function readTrazaCancelacion(cxpId) {
  try {
    cxpId = String(cxpId || '');
    if (!cxpId) return { ok: false, error: 'Falta cxpId' };
    var ss = _cxpCredSS();
    var out = { ok: true, creditos: [], devoluciones: [], aplicaciones: [] };

    // 1. Créditos generados por cancelar ESTA orden
    var credSh = ss.getSheetByName(CXP_CRED_TAB);
    var creditoIds = [];
    if (credSh) {
      var cd = credSh.getDataRange().getValues(); var ch = cd[0];
      for (var i = 1; i < cd.length; i++) {
        if (String(cd[i][_cxpColIdx(ch, 'CxPIdOrigen')]) !== cxpId) continue;
        var cid = String(cd[i][_cxpColIdx(ch, 'ID')]);
        creditoIds.push(cid);
        out.creditos.push({
          id: cid, monto: Number(cd[i][_cxpColIdx(ch, 'Monto')]) || 0,
          disponible: Number(cd[i][_cxpColIdx(ch, 'MontoDisponible')]) || 0,
          fecha: String(cd[i][_cxpColIdx(ch, 'Fecha')] || '')
        });
      }
    }

    // 2. Mapa de órdenes (Egresos) para nombrar la orden destino
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var egSh = ss.getSheetByName(egTab);
    var ordenInfo = {};
    if (egSh) {
      var ed = egSh.getDataRange().getValues();
      for (var e = 1; e < ed.length; e++) {
        var id = String(ed[e][0] || '');
        if (id) ordenInfo[id] = { proveedor: String(ed[e][4] || ''), concepto: String(ed[e][8] || ''), monto: parseFloat(ed[e][9]) || 0, fecha: (ed[e][1] instanceof Date) ? Utilities.formatDate(ed[e][1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(ed[e][1] || '') };
      }
    }

    // 3. Abonos que consumieron esos créditos → a qué orden se aplicaron,
    //    y qué OTROS abonos (depósito por la diferencia) tuvo esa orden
    var abSh = ss.getSheetByName(CXP_ABONO_TAB);
    if (abSh && creditoIds.length) {
      var ad = abSh.getDataRange().getValues(); var ah = ad[0];
      var destinos = {}; // cxpId destino → {aplicadoCredito, otrosAbonos:[]}
      for (var a = 1; a < ad.length; a++) {
        var reversado = ad[a][_cxpColIdx(ah, 'Reversado')] === true || String(ad[a][_cxpColIdx(ah, 'Reversado')]).toUpperCase() === 'TRUE';
        if (reversado) continue;
        var credId = String(ad[a][_cxpColIdx(ah, 'CreditoId')] || '');
        if (creditoIds.indexOf(credId) < 0) continue;
        var destCxp = String(ad[a][_cxpColIdx(ah, 'CxPId')] || '');
        if (!destinos[destCxp]) destinos[destCxp] = { cxpId: destCxp, aplicadoCredito: 0, otros: [] };
        destinos[destCxp].aplicadoCredito += Number(ad[a][_cxpColIdx(ah, 'Monto')]) || 0;
      }
      // Depósito/pago por la diferencia en cada orden destino
      Object.keys(destinos).forEach(function (destCxp) {
        for (var b = 1; b < ad.length; b++) {
          if (String(ad[b][_cxpColIdx(ah, 'CxPId')] || '') !== destCxp) continue;
          if (String(ad[b][_cxpColIdx(ah, 'Origen')] || '') !== 'pago') continue;
          var rev = ad[b][_cxpColIdx(ah, 'Reversado')] === true || String(ad[b][_cxpColIdx(ah, 'Reversado')]).toUpperCase() === 'TRUE';
          if (rev) continue;
          destinos[destCxp].otros.push({ monto: Number(ad[b][_cxpColIdx(ah, 'Monto')]) || 0, formaPago: String(ad[b][_cxpColIdx(ah, 'FormaPago')] || '') });
        }
        var info = ordenInfo[destCxp] || {};
        out.aplicaciones.push({
          cxpId: destCxp, proveedor: info.proveedor || '', concepto: info.concepto || '',
          montoOrden: info.monto || 0, fecha: info.fecha || '',
          aplicadoCredito: destinos[destCxp].aplicadoCredito,
          deposito: destinos[destCxp].otros
        });
      });
    }
    return out;
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── Cancelar una orden — el dinero ya pagado se reversa según el MODO
   que elija el usuario:
     modo 'credito' (default): queda como saldo a favor del proveedor
       (Creditos_Proveedor) — NO toca bancos ni conciliación; se aplica
       al siguiente documento de CxP desde el modal de Pagar.
     modo 'devolucion': el proveedor devolvió el dinero — se registra la
       ENTRADA en el banco/caja indicado (formaPago + fechaDevolucion)
       para que concilie con el depósito del estado de cuenta; NO se
       genera crédito.
   En ambos modos, los abonos que se habían pagado con un crédito previo
   regresan a ese crédito (ese dinero nunca salió del banco en ESTA
   orden). Si la orden no tenía pagos, solo se cancela. ─────────────── */
function cancelarOrdenCxP(body) {
  try {
    var ss = _cxpCredSS();
    var egTab = EGRESOS_TABS[2026] || 'Egresos2026';
    var sh = ss.getSheetByName(egTab);
    if (!sh) return { ok: false, error: 'Hoja Egresos no encontrada' };
    var rowNum = parseInt(body.rowNum);
    if (!rowNum || rowNum < 2) return { ok: false, error: 'Fila inválida' };

    var modo = body.modo === 'devolucion' ? 'devolucion' : 'credito';
    var formaPago = String(body.formaPago || '').trim();
    var fechaDev = String(body.fechaDevolucion || '').substring(0, 10) ||
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (modo === 'devolucion' && !formaPago)
      return { ok: false, error: 'Para registrar la devolución indica a qué banco/caja entró el dinero' };

    var estatusCol = _egColEnsure(sh, 'estatus', 'Estatus');
    var lastCol = Math.max(20, estatusCol);
    var rowData = sh.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var cxpId = String(rowData[0] || '');
    var proveedor = String(rowData[4] || '');
    var concepto = String(rowData[8] || '');
    var monto = parseFloat(rowData[9]) || 0;
    var pagado = rowData[13] === true || String(rowData[13]).toUpperCase() === 'TRUE';
    var estatusActual = String(rowData[estatusCol - 1] || '').trim();
    if (estatusActual === 'Cancelada') return { ok: false, error: 'Esta orden ya está cancelada' };

    var refDev = 'DEVOLUCIÓN ' + proveedor + ' — orden cancelada #' + cxpId + (concepto ? ' · ' + concepto : '');
    var creditoGenerado = 0, devolucionRegistrada = 0, creditoRestaurado = 0;

    function reversarDineroReal(montoRev, notaOrigen) {
      if (montoRev <= 0) return;
      if (modo === 'devolucion') {
        _cxpRutearDevolucionABanco(formaPago, montoRev, fechaDev, refDev);
        devolucionRegistrada += montoRev;
      } else {
        _crearCreditoProveedor({
          proveedor: proveedor, monto: montoRev,
          origen: notaOrigen, cxpIdOrigen: cxpId, usuario: body.usuario || ''
        });
        creditoGenerado += montoRev;
      }
    }

    var abonos = _cxpAbonosPorId(cxpId, true);
    abonos.forEach(function (ab) {
      if (ab.origen === 'credito' && ab.creditoId) {
        _restaurarCredito(ab.creditoId, ab.monto);
        creditoRestaurado += ab.monto;
      } else {
        reversarDineroReal(ab.monto, 'Orden cancelada #' + cxpId + ' — ' + concepto);
      }
      _marcarAbonoReversado(ab._rowNum);
    });
    // Pagada antes de que existiera este sistema: sin abonos registrados pero Pagado=TRUE
    if (pagado && abonos.length === 0 && monto > 0) {
      reversarDineroReal(monto, 'Orden cancelada #' + cxpId + ' — ' + concepto + ' (pago previo sin abono registrado)');
    }

    sh.getRange(rowNum, estatusCol).setValue('Cancelada');
    try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (e) {}
    var detalle = 'Cancelada [' + modo + ']'
      + (creditoGenerado ? ' | crédito generado: $' + creditoGenerado.toFixed(2) : '')
      + (devolucionRegistrada ? ' | devolución a ' + formaPago + ': $' + devolucionRegistrada.toFixed(2) : '')
      + (creditoRestaurado ? ' | crédito restaurado: $' + creditoRestaurado.toFixed(2) : '');
    logAudit(body.usuario || 'sistema', 'CxP', 'Cancelar', cxpId, 'Estatus', 'Activa', detalle);
    return { ok: true, modo: modo, creditoGenerado: creditoGenerado, devolucionRegistrada: devolucionRegistrada, creditoRestaurado: creditoRestaurado };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
