/* ══════════════════════════════════════════════════════════════════════════
   cancelacion.gs — CANCELACIÓN DE VENTAS (Ingresos) con traza completa
   ══════════════════════════════════════════════════════════════════════════

   PROBLEMA QUE RESUELVE
   ---------------------
   Hasta hoy la única forma de deshacer una venta era el bote de basura
   (deleteIngreso, finance.gs): BORRA las filas de BD_Ingresos y revierte los
   bancos. Para dinero eso es inaceptable: no queda registro de qué se vendió,
   ni de por qué se deshizo, ni de quién lo autorizó.

   Esta cancelación NO BORRA NADA. Marca la operación como Cancelada y guarda
   el expediente completo (motivo, evidencia, quién, cuándo, nº de operación
   bancaria e importe a devolver). La venta sigue visible en la tabla —
   tachada y con badge — porque ESE es el registro que el negocio necesita.

   MISMO PATRÓN CONCEPTUAL QUE LA CANCELACIÓN DE EGRESOS
   -----------------------------------------------------
   Reusa la lógica de cancelarOrdenCxP (cxp_creditos.gs:569), que el usuario ya
   conoce:
     · fila histórica que NO suma  → allá Estatus='Cancelada' + los lectores
                                     saltan (board.gs:47, semanal.gs:87);
                                     acá columna Cancelada=TRUE + los lectores
                                     de BD_Ingresos saltan (ver _ingEsCancelada).
     · columna append-safe         → allá _egColEnsure, acá _ingColEnsure.
     · idempotencia                → allá rechaza si ya está cancelada; acá
                                     devuelve ok:true + yaCancelada (para que
                                     se pueda reimprimir el comprobante).
     · traza del dinero            → allá crédito/devolución; acá el expediente
                                     de cancelación + cierre de CxCobrar.
     · logAudit al final.

   DECISIÓN — BANCOS: SÍ SE AFECTAN, COMO REVERSO (fila nueva, no borrado).
   -----------------------------------------------------------------------
   A Hestia el dinero se le va EN EL ACTO: Mercado Pago debita la devolución de
   la cuenta al procesarla. Que la paciente tarde días en ver el abono en su
   tarjeta es asunto del banco emisor — la cuenta de la clínica ya perdió ese
   dinero hoy. Si no se reflejara, el saldo de MP en el ERP mentiría y la
   conciliación no cuadraría contra el estado de cuenta real.

   Se modela como REVERSO, NO como borrado: la fila original del cobro (CARGO)
   queda intacta porque es historia real que sí ocurrió; se agrega una fila
   NUEVA que la neutraliza. Así la conciliación cuadra renglón por renglón
   contra el estado de cuenta de MP y queda traza de las dos cosas: que se
   cobró y que se devolvió.

   MP también devuelve la comisión (así lo reporta su detalle de devolución):
     original → Cobro +1800 | Comisiones  −73.08 | TotalVenta +1726.92
     reverso  → Cobro −1800 | Comisiones  +73.08 | TotalVenta −1726.92  Tipo=REVERSO
   El neto que baja de la cuenta es −1726.92 (lo que de verdad se había
   depositado), no −1800. Ver _canReversarBancos().

   Santander / Efectivo reusan _reverseIngresoBank (finance.gs), que ya sabe
   hacer el reverso por forma de pago.

   NUNCA se llama a updateIngresoConBancos / _reverseIngresoBankTodos: BORRAN y
   RE-CREAN filas de banco → destruyen la conciliación. Aquí se AGREGA.

   DECISIÓN — CxCOBRAR: SÍ SE CIERRA.
   ----------------------------------
   Una venta cancelada no puede seguir apareciendo como deuda del paciente. Se
   cierra reusando el mecanismo idempotente que ya existe
   (_cobRegistrarSaldoIngreso con saldo 0 → cobranza.gs:283, deja la fila con
   MontoCargo 0 y sale de Cuentas por Cobrar porque _cobReadCargos filtra
   `monto <= 0`). Encima se re-etiqueta el Estatus a 'Cancelado' para que el
   histórico no mienta diciendo 'Pagado'.

   PERMISO: cancelar_ingresos (propio, NO editar_ingresos). Ver cancelarVenta().

   DEPENDE DE (scope global de Apps Script, todo hoisted en runtime):
     finance.gs → _tokenHasPermission, _ingColEnsure, _ingEsCancelada,
                  _ingIdDeAnio, BD_INGRESOS_TAB, logAudit,
                  EGRESOS_DRIVE_PAGOS, _getOrCreateMonthFolder
     cobranza.gs → _cobRegistrarSaldoIngreso (opcional, con guarda typeof)
   ══════════════════════════════════════════════════════════════════════════ */

var CANCEL_VER = '2026.07.15.1';

/* Encabezados de las 2 columnas nuevas de BD_Ingresos. Se crean al vuelo con
   _ingColEnsure (append al final, jamás por posición fija).
   OJO con la búsqueda por substring de _ingColEnsure: 'cancelada' NO hace match
   dentro de 'cancelaciondata' ni al revés → las dos columnas conviven bien. */
var CANCEL_COL_FLAG = 'Cancelada';        // TRUE / vacío
var CANCEL_COL_DATA = 'CancelacionData';  // JSON con el expediente completo

function _canNum(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v || '').replace(/[$,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function _canHoy() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd HH:mm');
}

/* Sube la evidencia opcional a Drive. Reusa la MISMA carpeta y el mismo patrón
   de uploadFile (finance.gs:5190): Contabilidad\Pagos\<Año>\<Mes>, compartido
   por liga. Nunca debe tumbar la cancelación: si Drive falla, se cancela igual
   y se reporta el fallo. */
function _canSubirEvidencia(op, body) {
  try {
    if (!body || !body.evidenciaBase64) return { url: '', error: '' };
    var parentId = (typeof EGRESOS_DRIVE_PAGOS !== 'undefined') ? EGRESOS_DRIVE_PAGOS : '';
    if (!parentId) return { url: '', error: 'Carpeta de Drive no configurada' };
    var hoy = new Date();
    var folder = (typeof _getOrCreateMonthFolder === 'function')
      ? _getOrCreateMonthFolder(parentId, hoy.getFullYear(), hoy.getMonth())
      : DriveApp.getFolderById(parentId);
    var limpio = String(body.evidenciaNombre || 'evidencia.pdf').replace(/[\\\/:*?"<>|]/g, '_');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(body.evidenciaBase64),
      body.evidenciaMime || 'application/octet-stream',
      'CANCELACION_' + op + '_' + limpio
    );
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (_sh) {}
    return { url: file.getUrl(), error: '' };
  } catch (ex) {
    return { url: '', error: ex.message };
  }
}

/* Cierra la cuenta por cobrar de una OP cancelada. Reusa el upsert idempotente
   de cobranza.gs; si cobranza.gs no está desplegado, no pasa nada. */
function _canCerrarCxCobrar(op, paciente, categoria, fecha) {
  try {
    if (typeof _cobRegistrarSaldoIngreso !== 'function') return { cerrada: false, motivo: 'cobranza.gs no disponible' };
    // saldo 0 → la fila auto-generada queda en MontoCargo 0 y sale de Cuentas por Cobrar.
    // El upsert AVISA si no encontró renglón auto de esta OP ('sin-cambio'): entonces no
    // cerró nada y decir lo contrario sería mentir. Pasa con un cargo legacy cuya OP se
    // llenó a mano (sin la Nota 'auto-ingreso' que el upsert usa para localizarlo), y con
    // la venta de contado que nunca tuvo deuda. Se reporta tal cual y el caller decide.
    var _up = _cobRegistrarSaldoIngreso(op, paciente, categoria, 0, fecha);
    if (_up && _up.ok === false)
      return { cerrada: false, motivo: 'no se pudo cerrar la cuenta por cobrar: ' + (_up.error || 'error') };
    if (_up && _up.accion === 'sin-cambio')
      return { cerrada: false, motivo: 'la venta no tenía cuenta por cobrar auto-generada que cerrar '
        + '(venta de contado, o cargo capturado a mano sin la marca auto-ingreso: revísalo en Cobranza)' };
    // Re-etiquetar el estatus: 'Pagado' (lo que pone el upsert) mentiría en el histórico.
    try {
      var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
      var sh = ss.getSheetByName(typeof COBRANZA_CARGOS !== 'undefined' ? COBRANZA_CARGOS : 'Cuentas_Cobrar');
      if (sh) {
        var raw = sh.getDataRange().getValues();
        var hdr = (raw[0] || []).map(function (x) { return String(x).trim().toLowerCase(); });
        var iOp = hdr.indexOf('op');      if (iOp < 0) iOp = 1;
        var iEst = hdr.indexOf('estatus'); if (iEst < 0) iEst = 6;
        var iNota = hdr.indexOf('nota');   if (iNota < 0) iNota = 7;
        for (var r = 1; r < raw.length; r++) {
          if (String(raw[r][iOp] || '').trim() !== String(op).trim()) continue;
          if (String(raw[r][iNota] || '').toLowerCase().indexOf('auto-ingreso') < 0) continue;
          sh.getRange(r + 1, iEst + 1).setValue('Cancelado');
          break;
        }
      }
    } catch (_lbl) { /* el re-etiquetado es cosmético: nunca debe tumbar la cancelación */ }
    return { cerrada: true, motivo: '' };
  } catch (ex) {
    return { cerrada: false, motivo: ex.message };
  }
}

/* ═════════ REVERSO BANCARIO ═════════════════════════════════════════════
   Se reversa lo que DE VERDAD está registrado en el banco para esta OP (no lo
   que "debería" estar según BD_Ingresos): así el reverso neutraliza exactamente
   los renglones que existen y la conciliación cuadra uno a uno.

   Amarre por etiqueta en Observaciones. Conviven DOS formatos históricos:
     '[' + opId + ']'        → canónico (finance.gs:6013, _bankDeleteByOp)
     '[OP #' + opId + ']'    → el que escribe _abonoRutearABanco (finance.gs:5142)
   Se buscan los dos. No se solapan: '[OP #OP-1]' no contiene '[OP-1]'.
   ═══════════════════════════════════════════════════════════════════════ */
function _canTagMatch(obs, opId) {
  var s = String(obs || '');
  return s.indexOf('[' + opId + ']') > -1 || s.indexOf('[OP #' + opId + ']') > -1;
}
/* ¿Esta fila YA es un reverso/cancelación? (2ª red de idempotencia, además del
   flag Cancelada en BD_Ingresos). */
function _canEsFilaReverso(obs, tipo) {
  var s = String(obs || '').toUpperCase();
  return String(tipo || '').trim().toUpperCase() === 'REVERSO'
      || s.indexOf('REVERSO') > -1 || s.indexOf('CANCELACIÓN') > -1 || s.indexOf('CANCELACION') > -1;
}

/* Genera el/los movimiento(s) de reverso de una OP. Devuelve el detalle para
   el log y el reporte. Nunca debe tumbar la cancelación: si el banco falla, la
   venta igual queda cancelada y se reporta el error. */
function _canReversarBancos(opId, pac, fechaCancel) {
  var out = { movimientos: [], yaReversado: false, error: '' };
  try {
    var obs = 'Cancelación · Px. ' + String(pac || '') + ' [' + opId + ']';

    /* ── Mercado Pago (incluye TDC/TDD/AMEX/Transferencia: todo rutea a MP) ── */
    try {
      var shMP = (typeof _bankSheetByKey === 'function') ? _bankSheetByKey('mercadopago') : null;
      if (shMP && shMP.getLastRow() > 1) {
        var lastC = Math.max(10, shMP.getLastColumn());
        var vals = shMP.getRange(2, 1, shMP.getLastRow() - 1, lastC).getValues();
        var cobro = 0, com = 0, hay = false;
        for (var r = 0; r < vals.length; r++) {
          if (!_canTagMatch(vals[r][8], opId)) continue;          // col I (idx 8) = Observaciones
          if (_canEsFilaReverso(vals[r][8], vals[r][9])) { out.yaReversado = true; continue; } // col J (idx 9) = Tipo
          cobro += _canNum(vals[r][2]);   // col C = Cobro (positivo)
          com += _canNum(vals[r][3]);     // col D = Comisiones (negativo)
          hay = true;
        }
        if (hay && !out.yaReversado && (cobro !== 0 || com !== 0)) {
          var mes = String(fechaCancel || '').substring(0, 7);
          // Cobro −1800 | Comisiones +73.08 | TotalVenta −1726.92 | Tipo REVERSO.
          // saveBankRow calcula solo el %Comisión (idx 4) y el Saldo corrido (idx 6): NO se duplican.
          saveBankRow('mercadopago', [mes, fechaCancel, -cobro, -com, 0, -(cobro + com), 0, false, obs, 'REVERSO']);
          out.movimientos.push({ banco: 'mercadopago', cobro: -cobro, comision: -com, neto: -(cobro + com) });
        }
      }
    } catch (eMP) { out.error += 'MP: ' + eMP.message + ' '; }

    /* ── Santander: el reverso es un RETIRO por lo depositado ── */
    try {
      var shSA = (typeof _bankSheetByKey === 'function') ? _bankSheetByKey('santander') : null;
      if (shSA && shSA.getLastRow() > 1) {
        var vs = shSA.getRange(2, 1, shSA.getLastRow() - 1, Math.max(5, shSA.getLastColumn())).getValues();
        var dep = 0, hayS = false, yaS = false;
        for (var s2 = 0; s2 < vs.length; s2++) {
          if (!_canTagMatch(vs[s2][4], opId)) continue;           // col E (idx 4) = Referencia/Obs
          if (_canEsFilaReverso(vs[s2][4], '')) { yaS = true; continue; }
          dep += _canNum(vs[s2][1]);      // col B = Depósito
          hayS = true;
        }
        if (hayS && !yaS && dep > 0) {
          // Reusa el reverso por forma de pago que ya existe (finance.gs).
          _reverseIngresoBank(opId, 'Santander', fechaCancel, dep, obs);
          out.movimientos.push({ banco: 'santander', retiro: dep });
        } else if (yaS) { out.yaReversado = true; }
      }
    } catch (eSA) { out.error += 'Santander: ' + eSA.message + ' '; }

    /* ── Efectivo / Caja Chica: el reverso es una SALIDA ── */
    try {
      if (typeof getCajaChicaSheet === 'function') {
        var shCC = getCajaChicaSheet();
        if (shCC && shCC.getLastRow() > 1) {
          var dcc = shCC.getDataRange().getValues();
          var hcc = (dcc[0] || []).map(function (h) { return String(h).trim().toUpperCase(); });
          var iC = hcc.indexOf('CONCEPTO'), iE = hcc.indexOf('ENTRADA');
          if (iC > -1 && iE > -1) {
            var ent = 0, hayC = false, yaC = false;
            for (var c2 = 1; c2 < dcc.length; c2++) {
              if (!_canTagMatch(dcc[c2][iC], opId)) continue;
              if (_canEsFilaReverso(dcc[c2][iC], '')) { yaC = true; continue; }
              ent += _canNum(dcc[c2][iE]);
              hayC = true;
            }
            if (hayC && !yaC && ent > 0) {
              _reverseIngresoBank(opId, 'Efectivo', fechaCancel, ent, obs);
              out.movimientos.push({ banco: 'cajachica', salida: ent });
            } else if (yaC) { out.yaReversado = true; }
          }
        }
      }
    } catch (eCC) { out.error += 'Caja chica: ' + eCC.message + ' '; }

  } catch (ex) { out.error += ex.message; }
  return out;
}

/* ─────────────────────────────────────────────────────────────────────────
   cancelarVenta(body)
   ─────────────────────────────────────────────────────────────────────────
   body: { token, usuario, op, anio?, motivo*, opBanco, importeDevolver,
           medioPago, evidenciaBase64?, evidenciaNombre?, evidenciaMime? }

   PERMISO: 'cancelar_ingresos' — PROPIO, no reusa editar_ingresos.
   ¿Por qué?
     1) Cancelar es más grave que editar: mueve dinero hacia afuera (dispara un
        reembolso al paciente) y retira la venta de todos los reportes.
        editar_ingresos hoy lo tienen gerente/capturista para corregir typos.
     2) Un permiso NUEVO tiene el failure-mode correcto: ningún rol lo trae
        todavía, así que al desplegar SOLO admin/director pueden cancelar
        (_tokenHasPermission devuelve true para admin/director, finance.gs:4849)
        hasta que el administrador lo otorgue explícitamente en Panel de Control.
        Reusar editar_ingresos habría dado el poder de cancelar, en silencio y
        de inmediato, a todo el que hoy puede corregir una fecha.
   ───────────────────────────────────────────────────────────────────────── */
function cancelarVenta(body) {
  try {
    body = body || {};
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token || '', 'cancelar_ingresos'))
      return { ok: false, error: 'Sin autorización para cancelar ventas. Solicita al administrador el permiso cancelar_ingresos.', version: CANCEL_VER };

    var op = String(body.op || '').trim();
    if (!op) return { ok: false, error: 'OP inválida', version: CANCEL_VER };

    // MOTIVO OBLIGATORIO: sin esto la cancelación no vale como registro.
    var motivo = String(body.motivo || '').trim();
    if (motivo.length < 5)
      return { ok: false, error: 'El motivo de la cancelación es obligatorio (mínimo 5 caracteres).', version: CANCEL_VER };

    var opBanco = String(body.opBanco || '').trim();
    var medioPago = String(body.medioPago || '').trim();
    var usuario = String(body.usuario || '').trim() || 'sistema';
    var anio = parseInt(body.anio, 10) || new Date().getFullYear();

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado, reintenta en un momento.', version: CANCEL_VER };
    try {
      var ss = SpreadsheetApp.openById(_ingIdDeAnio(anio));
      var sh = ss.getSheetByName(BD_INGRESOS_TAB);
      if (!sh) return { ok: false, error: 'No se encontró BD_Ingresos', version: CANCEL_VER };

      // Columnas append-safe. Se resuelven ANTES de leer los datos porque
      // _ingColEnsure puede agregar columnas (y con ello cambiar el ancho).
      var cFlag = _ingColEnsure(sh, 'cancelada', CANCEL_COL_FLAG);
      var cData = _ingColEnsure(sh, 'cancelaciondata', CANCEL_COL_DATA);

      var data = sh.getDataRange().getValues();
      var H = (data[0] || []).map(function (x) { return String(x || '').trim().toLowerCase(); });
      function hc() {
        for (var a = 0; a < arguments.length; a++) { var k = H.indexOf(arguments[a]); if (k > -1) return k; }
        return -1;
      }
      var iOp = hc('op'); if (iOp < 0) iOp = 0;
      var iFecha = hc('fecha'), iPac = hc('paciente'), iCat = hc('categoria', 'categoría');
      var iProd = hc('producto'), iFP = hc('formapago', 'forma de pago');
      var iTot = hc('totalpagar', 'total a pagar', 'total'), iPag = hc('pagado');

      var rowNums = [], totalOP = 0, pagadoOP = 0;
      var pac = '', cat = '', fecha = '', fp = '', productos = [];
      var yaCanceladaJson = '';
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][iOp] || '').trim() !== op) continue;
        rowNums.push(r + 1);
        totalOP += _canNum(data[r][iTot]);
        pagadoOP += _canNum(data[r][iPag]);
        if (!pac && iPac > -1) pac = String(data[r][iPac] || '');
        if (!cat && iCat > -1) cat = String(data[r][iCat] || '');
        if (!fp && iFP > -1) fp = String(data[r][iFP] || '');
        if (!fecha && iFecha > -1) {
          var fv = data[r][iFecha];
          fecha = (fv instanceof Date)
            ? Utilities.formatDate(fv, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd')
            : String(fv || '');
        }
        if (iProd > -1 && data[r][iProd]) productos.push(String(data[r][iProd]));
        if (typeof _ingEsCancelada === 'function' && _ingEsCancelada(data[r][cFlag - 1]) && !yaCanceladaJson)
          yaCanceladaJson = String(data[r][cData - 1] || '');
      }
      if (!rowNums.length) return { ok: false, error: 'OP ' + op + ' no encontrada en BD_Ingresos', version: CANCEL_VER };

      // IDEMPOTENCIA: ya cancelada → no se re-escribe (el expediente original es
      // inmutable) y NO se vuelve a tocar CxCobrar. Se devuelve ok:true con los
      // datos guardados para que el frontend pueda reimprimir el comprobante.
      if (yaCanceladaJson) {
        var prev = {};
        try { prev = JSON.parse(yaCanceladaJson) || {}; } catch (_pj) {}
        return {
          ok: true, yaCancelada: true, op: op, version: CANCEL_VER,
          cancelacion: prev, lineas: rowNums.length,
          mensaje: 'Esta venta ya estaba cancelada el ' + (prev.fechaHora || '?') + ' por ' + (prev.usuario || '?') + '.'
        };
      }

      // Importe a devolver: lo confirma el usuario. Si no lo manda, se asume lo
      // efectivamente pagado por la paciente (no el total facturado).
      var importe = (body.importeDevolver === '' || body.importeDevolver == null)
        ? pagadoOP : Math.max(0, _canNum(body.importeDevolver));

      var ev = _canSubirEvidencia(op, body);

      // BANCOS: reverso ANTES de marcar la fila, para que si el banco truena la
      // venta no quede marcada como cancelada sin su movimiento (mejor reintentar).
      var fechaCancel = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd');
      var bancos = _canReversarBancos(op, pac, fechaCancel);

      var exp = {
        motivo: motivo,
        usuario: usuario,
        fechaHora: _canHoy(),
        opBanco: opBanco,
        importeDevolver: importe,
        medioPago: medioPago || fp,
        evidenciaUrl: ev.url || '',
        paciente: pac,
        fechaVenta: fecha,
        totalOP: totalOP,
        pagadoOP: pagadoOP,
        bancos: bancos.movimientos
      };
      var expJson = JSON.stringify(exp);

      // NO se borra la fila: se marca. Todas las líneas de la OP quedan marcadas
      // para que cualquier lector, filtre por la línea que filtre, la vea cancelada.
      for (var k = 0; k < rowNums.length; k++) {
        sh.getRange(rowNums[k], cFlag).setValue(true);
        sh.getRange(rowNums[k], cData).setValue(expJson);
      }

      // CxCobrar: la venta ya no existe → no puede seguir siendo deuda.
      var cxc = _canCerrarCxCobrar(op, pac, cat, fecha);

      try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (_c) {}

      try {
        logAudit(usuario, 'Ingresos', 'Cancelar', op, 'Cancelada', 'Activa',
          'motivo: ' + motivo
          + ' | devolver: $' + importe.toFixed(2)
          + (opBanco ? ' | op. bancaria: ' + opBanco : '')
          + (ev.url ? ' | evidencia: ' + ev.url : ' | sin evidencia')
          + ' | ' + rowNums.length + ' lineas'
          + (cxc.cerrada ? ' | CxCobrar cerrada' : ' | CxCobrar NO cerrada: ' + (cxc.motivo || 'sin motivo'))
          + (bancos.movimientos.length
              ? ' | reverso banco: ' + bancos.movimientos.map(function (m) { return m.banco; }).join(',')
              : ' | sin reverso de banco'));
      } catch (_a) {}

      return {
        ok: true, op: op, version: CANCEL_VER, lineas: rowNums.length,
        // cxCobrarCerrada viaja SIN maquillar: false + motivo cuando no se cerró nada.
        // La cancelación en sí es válida igual (la venta ya quedó marcada y el banco
        // reversado); lo que no se puede es afirmar que se cerró una deuda que no se tocó.
        cancelacion: exp, cxCobrarCerrada: cxc.cerrada, cxCobrarMotivo: cxc.motivo || '',
        evidenciaError: ev.error || '',
        bancos: bancos.movimientos, bancosError: bancos.error || '',
        productos: productos
      };
    } finally {
      try { lock.releaseLock(); } catch (_rl) {}
    }
  } catch (ex) {
    return { ok: false, error: ex.message, version: CANCEL_VER };
  }
}
