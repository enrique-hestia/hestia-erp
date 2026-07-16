/* ═══════════════════════════════════════════════════════════════════════════
   AUDITORÍA DE EGRESOS SIN MOVIMIENTO BANCARIO  ·  SOLO LECTURA
   ───────────────────────────────────────────────────────────────────────────
   Por qué existe: durante meses, _egresoResyncBanco BORRABA la fila del banco
   de forma incondicional y solo la re-creaba si la forma de pago caía en una
   lista cerrada ('Santander' | 'Transferencia' | 'AMEX' | 'Mercado Pago' |
   'TDC' | 'TDD' | 'Efectivo', con match EXACTO). Como el dropdown de formas de
   pago es editable por el usuario, cualquier valor fuera de esa lista — o el
   mismo valor con otra caja ('santander') — producía: 1 fila borrada, 0 creadas,
   cero avisos. El egreso seguía existiendo; el movimiento del banco, no; y el
   saldo del banco quedaba mal por el monto del egreso.

   El bug ya está arreglado (plan-antes-de-borrar + error duro), pero los egresos
   huérfanos que dejó siguen ahí. Esta función los ENCUENTRA. No escribe nada:
   ni una celda, ni una fila. Es un reporte.

   Uso desde el editor de Apps Script:
       auditarEgresosSinBanco(2026)   → Logger.log del reporte + objeto de vuelta
   ═══════════════════════════════════════════════════════════════════════════ */
function auditarEgresosSinBanco(anio) {
  anio = parseInt(anio, 10) || new Date().getFullYear();
  try {
    var ss = SpreadsheetApp.openById(_egIdDeAnio(anio));
    var sheet = ss.getSheetByName(_egTabDeAnio(anio));
    if (!sheet) return { ok: false, error: 'No existe la pestaña ' + _egTabDeAnio(anio) };

    var raw = sheet.getDataRange().getValues();
    if (raw.length < 2) return { ok: true, anio: anio, total: 0, huerfanos: [], impacto: {} };

    var hdrs = raw[0].map(function (h) {
      return String(h).trim().toLowerCase()
        .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
        .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u');
    });
    function col(name) { for (var c = 0; c < hdrs.length; c++) if (hdrs[c].indexOf(name) > -1) return c; return -1; }
    var iFecha = col('fecha'), iProv = col('proveedor'), iConc = col('concepto'),
        iMonto = col('egresos'), iPagado = col('pagado'), iForma = col('forma'),
        iEstatus = col('estatus');

    // ── Índice de etiquetas [Egreso #id] presentes HOY en los 4 destinos ──
    var tags = _auditRecolectarTagsBanco();

    var huerfanos = [], impacto = { santander: 0, amex: 0, mercadopago: 0, cajachica: 0, sinDestino: 0 };
    var revisados = 0;

    for (var r = 1; r < raw.length; r++) {
      var row = raw[r];
      var egId = String(row[0] || '').trim();
      if (!egId) continue;
      var estatus = iEstatus > -1 ? String(row[iEstatus] || '').trim().toLowerCase() : '';
      if (estatus.indexOf('cancel') > -1) continue;          // cancelado: no debe tener fila

      var monto = iMonto > -1 ? (parseFloat(String(row[iMonto] || '').replace(/[$,\s]/g, '')) || 0) : 0;
      var pgVal = iPagado > -1 ? row[iPagado] : false;
      var pagado = (pgVal === true || String(pgVal).toLowerCase() === 'true' || String(pgVal).toUpperCase() === 'VERDADERO');
      if (!pagado || monto <= 0) continue;                    // sin dinero movido, no aplica
      revisados++;

      if (_auditTieneFila(tags, egId)) continue;              // tiene su fila: sano

      // ── Huérfano: pagado, con monto, y sin fila en ningún banco ──
      var formaPago = iForma > -1 ? String(row[iForma] || '').trim() : '';
      var ruta = (typeof _egRutaBanco === 'function') ? _egRutaBanco(formaPago) : null;

      var motivo, destino;
      if (!formaPago) {
        motivo = 'Sin forma de pago: nunca se supo de qué banco salió.';
        destino = 'sinDestino';
      } else if (!ruta) {
        motivo = 'Forma de pago "' + formaPago + '" fuera del ruteo conocido → el resync borraba y no recreaba.';
        destino = 'sinDestino';
      } else {
        motivo = 'Forma de pago "' + formaPago + '" SÍ rutea a ' + ruta.banco +
                 ' pero la fila no está. Causas probables: se borró en un resync fallido, ' +
                 'abonos raros (todos a crédito / sin forma de pago), o se borró a mano.';
        destino = ruta.banco;
      }

      var fRaw = iFecha > -1 ? row[iFecha] : '';
      var fecha = (fRaw instanceof Date)
        ? Utilities.formatDate(fRaw, 'America/Mexico_City', 'yyyy-MM-dd')
        : String(fRaw || '').substring(0, 10);

      impacto[destino] = (impacto[destino] || 0) + monto;
      huerfanos.push({
        id: egId, fila: r + 1, fecha: fecha,
        proveedor: iProv > -1 ? String(row[iProv] || '') : '',
        concepto: iConc > -1 ? String(row[iConc] || '') : '',
        monto: monto, formaPago: formaPago || '(vacía)',
        destinoEsperado: destino, motivo: motivo,
        reparable: !!ruta
      });
    }

    var rep = ['', '═══ EGRESOS PAGADOS SIN MOVIMIENTO EN BANCO · ' + anio + ' ═══',
               'Egresos pagados revisados: ' + revisados,
               'HUÉRFANOS ENCONTRADOS: ' + huerfanos.length, ''];
    huerfanos.forEach(function (h) {
      rep.push('  #' + h.id + '  ' + h.fecha + '  $' + h.monto.toFixed(2) + '  [' + h.formaPago + ']  ' +
               h.proveedor + ' · ' + h.concepto);
      rep.push('      → ' + h.motivo + (h.reparable ? '  (REPARABLE)' : '  (NO reparable: ruteo desconocido)'));
    });
    rep.push('', '── Impacto en el saldo de cada banco (monto que le FALTA descontar) ──');
    Object.keys(impacto).forEach(function (k) {
      if (impacto[k]) rep.push('  ' + k + ': $' + impacto[k].toFixed(2));
    });
    if (!huerfanos.length) rep.push('  Ninguno. Los saldos cuadran con los egresos pagados.');
    rep.push('');
    var texto = rep.join('\n');
    Logger.log(texto);

    return { ok: true, anio: anio, revisados: revisados, total: huerfanos.length,
             huerfanos: huerfanos, impacto: impacto, reporte: texto };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}

/* Recolecta TODAS las observaciones/conceptos de los 4 destinos de dinero.
   Solo getValues(): ni un setValue en todo el archivo. */
function _auditRecolectarTagsBanco() {
  var out = [];
  ['santander', 'amex', 'mercadopago'].forEach(function (key) {
    try {
      var sh = _bankSheetByKey(key); if (!sh) return;
      var lr = sh.getLastRow(); if (lr < 2) return;
      var idx = _egBankObsIdx(key);
      var vals = sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues();
      for (var i = 0; i < vals.length; i++) out.push(vals[i][idx]);
    } catch (e) { Logger.log('_auditRecolectarTagsBanco ' + key + ': ' + e.message); }
  });
  try {
    var cc = getCajaChicaSheet();
    if (cc) {
      var d = cc.getDataRange().getValues();
      if (d.length > 1) {
        var ic = d[0].map(function (h) { return String(h).trim().toUpperCase(); }).indexOf('CONCEPTO');
        if (ic > -1) for (var r = 1; r < d.length; r++) out.push(d[r][ic]);
      }
    }
  } catch (e) { Logger.log('_auditRecolectarTagsBanco caja: ' + e.message); }
  return out;
}

function _auditTieneFila(tags, egId) {
  for (var i = 0; i < tags.length; i++) if (_egBankTagMatch(tags[i], egId)) return true;
  return false;
}

/* ── Reparación asistida. NO tiene botón y NO corre por accidente: exige
   {aplicar:true} Y una lista EXPLÍCITA de ids. Sin ambas cosas hace dry-run.
   Idempotente (usa el mismo resync, que borra-y-recrea). Rechaza — no adivina —
   los egresos cuya forma de pago no está en el ruteo. ── */
function repararEgresosSinBanco(body) {
  body = body || {};
  var anio = parseInt(body.anio, 10) || new Date().getFullYear();
  var ids = Array.isArray(body.ids) ? body.ids.map(function (x) { return String(x).trim(); }).filter(String) : [];
  var aplicar = (body.aplicar === true);

  if (!ids.length) return { ok: false, error: 'Faltan ids: hay que pasar la lista EXPLÍCITA de egresos a reparar (los da auditarEgresosSinBanco).' };

  var aud = auditarEgresosSinBanco(anio);
  if (!aud.ok) return aud;
  var porId = {}; aud.huerfanos.forEach(function (h) { porId[h.id] = h; });

  var plan = [], rechazados = [];
  ids.forEach(function (id) {
    var h = porId[id];
    if (!h) { rechazados.push(id + ': no aparece como huérfano en ' + anio + ' (¿ya se reparó?)'); return; }
    if (!h.reparable) { rechazados.push(id + ': forma de pago "' + h.formaPago + '" fuera del ruteo. Corrige el egreso primero — no voy a adivinar el banco.'); return; }
    plan.push(h);
  });

  if (!aplicar) {
    return { ok: true, dryRun: true, anio: anio,
             mensaje: 'Dry-run: NO se escribió nada. Vuelve a llamar con aplicar:true para ejecutar.',
             repararia: plan, rechazados: rechazados };
  }

  var ss = SpreadsheetApp.openById(_egIdDeAnio(anio));
  var sheet = ss.getSheetByName(_egTabDeAnio(anio));
  var hechos = [], errores = [];
  plan.forEach(function (h) {
    try {
      _egresoResyncBanco(sheet, h.fila);        // recrea la fila que falta
      logAudit(body.usuario || 'sistema', 'Egresos', 'Reparar banco huerfano', 'Egreso #' + h.id,
        h.proveedor + ' · ' + h.concepto, 'sin fila', h.destinoEsperado + ' $' + h.monto);
      hechos.push(h.id);
    } catch (e) { errores.push(h.id + ': ' + e.message); }
  });
  return { ok: true, anio: anio, reparados: hechos, errores: errores, rechazados: rechazados };
}
