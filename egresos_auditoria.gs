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

   ── LA LECCIÓN, aprendida en producción (2026-07-16) ────────────────────────
   La primera versión juzgaba SOLO por la etiqueta [Egreso #id] y reportó
   628 huérfanos de 676 (93%). Era MENTIRA: se comprobaron 5 a mano y los 5
   existían en el banco. "Reparar" habría duplicado $11,919,621.56.

   El error: la etiqueta la estampa ÚNICAMENTE el ERP al pagar (finance.gs:1896).
   Los movimientos de enero–mayo 2026 son carga histórica — reales, correctos, y
   sin etiqueta porque nunca pasaron por el ERP. La auditoría no medía "falta el
   movimiento": medía "esto no se pagó desde el ERP". Dato duro: hay ~982
   movimientos reales en ene–may y CERO etiquetas entre ellos; la etiqueta más
   antigua del libro es del 2026-06-30.

   Ahora, antes de acusar a un egreso hacen falta TRES cosas:
     1. Que no tenga etiqueta.
     2. Que sea POSTERIOR a la primera etiqueta del libro (_auditEraEtiquetas,
        que se calcula sola). Antes de esa fecha no hay nada que juzgar.
     3. Que NO exista un movimiento suyo por fecha ±2d + monto — incluida la
        variante N:1, porque un egreso puede salir como VARIAS filas que suman
        (un AMEX de $76,737.31 fueron dos cargos: 49,900 + 26,837.31).

   Regla de la casa: ante la duda, callar. Un huérfano no reportado cuesta una
   revisión manual; un falso positivo "reparado" duplica dinero en el banco.

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

    // ── Movimientos de salida que existen HOY en los 4 destinos ──
    var idx = _auditIndiceBanco();
    var eraDesde = _auditEraEtiquetas(idx);   // '' = el ERP nunca ha etiquetado nada

    var huerfanos = [], impacto = { santander: 0, amex: 0, mercadopago: 0, cajachica: 0, sinDestino: 0 };
    var revisados = 0, sanosEtiqueta = 0, noEvaluables = 0, sinEtiquetaConMov = 0;

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

      if (_auditTieneEtiqueta(idx, egId)) { sanosEtiqueta++; continue; }   // tiene su fila: sano

      var fRaw0 = iFecha > -1 ? row[iFecha] : '';
      var fEg = (fRaw0 instanceof Date)
        ? Utilities.formatDate(fRaw0, 'America/Mexico_City', 'yyyy-MM-dd')
        : String(fRaw0 || '').substring(0, 10);

      // ── Sin etiqueta NO significa sin movimiento ────────────────────────────
      // La etiqueta [Egreso #id] solo la estampa el ERP al pagar (finance.gs:1896).
      // Todo lo anterior a `eraDesde` es carga histórica: el movimiento existe en
      // el banco y nunca llevó etiqueta. Juzgarlo por etiqueta declaraba huérfano
      // el 93% del libro (628 de 676) y "repararlo" habría duplicado $11.9M.
      //
      // Se descarta —sin acusar— en tres casos, todos "no evaluable", no "sano":
      //   · !eraDesde: no hay NI UNA etiqueta en los bancos → el ERP nunca pagó
      //     nada → no existe forma de distinguir un huérfano de una carga vieja.
      //   · !fEg: egreso sin fecha → no hay con qué cruzar contra el banco.
      //   · fEg < eraDesde: carga histórica.
      // El objetivo de esta auditoría son las filas que el resync se comió, y esas
      // son por definición de la era del ERP. Fuera de ahí, callar es lo correcto.
      if (!eraDesde || !fEg || fEg < eraDesde) { noEvaluables++; continue; }

      var formaPago = iForma > -1 ? String(row[iForma] || '').trim() : '';
      var ruta = (typeof _egRutaBanco === 'function') ? _egRutaBanco(formaPago) : null;

      // Aun dentro de la era de las etiquetas: antes de acusar, se busca el
      // movimiento por FECHA + MONTO. Si está, no hay nada que reparar — y
      // recrearlo sería duplicar dinero.
      if (ruta && idx[ruta.banco]) {
        var mv = _auditBuscaMovimiento(idx[ruta.banco], fEg, monto);
        if (mv.hit) { sinEtiquetaConMov++; continue; }
      }

      // ── Huérfano: pagado, con monto, sin etiqueta Y sin movimiento que le case ──
      var motivo, destino;
      if (!formaPago) {
        motivo = 'Sin forma de pago: nunca se supo de qué banco salió.';
        destino = 'sinDestino';
      } else if (!ruta) {
        motivo = 'Forma de pago "' + formaPago + '" fuera del ruteo conocido → el resync borraba y no recreaba.';
        destino = 'sinDestino';
      } else {
        motivo = 'Forma de pago "' + formaPago + '" rutea a ' + ruta.banco + ', pero ahí no hay etiqueta ' +
                 'NI ningún movimiento de $' + monto.toFixed(2) + ' entre el ' + fEg + ' ±2 días. ' +
                 'Candidato real: pudo borrarse en un resync fallido o a mano. VERIFÍCALO en el banco antes de reparar.';
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
               'Egresos pagados revisados: ' + revisados, '',
               '── Cómo se descartaron (esto NO son problemas) ──',
               '  Con su etiqueta [Egreso #id]:            ' + sanosEtiqueta,
               '  No evaluables (carga histórica, sin fecha): ' + noEvaluables +
                 '   ← el ERP aún no etiquetaba; su movimiento existe sin etiqueta',
               '  Sin etiqueta pero CON su movimiento:     ' + sinEtiquetaConMov +
                 '   ← casado por fecha ±2d + monto',
               '',
               'CANDIDATOS A HUÉRFANO: ' + huerfanos.length,
               (huerfanos.length ? '  ⚠ Verifica CADA UNO en el banco antes de reparar: reparar algo que sí existe DUPLICA el dinero.' : ''),
               ''];
    if (!eraDesde) {
      rep.push('  ⚠ No hay ni una etiqueta [Egreso #] en los bancos: el ERP nunca ha pagado nada,',
               '    así que no hay forma de distinguir un huérfano de una carga histórica.',
               '    No se acusa a nadie.', '');
    }
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

/* ── LAYOUT REAL DE CADA DESTINO ──────────────────────────────────────────────
   Verificado contra el .xlsx exportado (2026-07-16). No se puede deducir del
   código ni asumir que los tres bancos se parecen: NO se parecen.
     · La FECHA no está en el mismo índice (MP la tiene en la 1, no en la 0).
     · MP guarda la fecha CON HORA (serial 45888.64375) → comparar Date crudo
       nunca casa; hay que normalizar a yyyy-MM-dd.
     · El SIGNO del egreso cambia por banco: Santander/AMEX positivo, MP negativo.
     · En AMEX un Movimiento NEGATIVO es un PAGO a la tarjeta, no un gasto.
   `obs` coincide con _egBankObsIdx (finance.gs:3028); se repite aquí para que
   este archivo se lea completo sin saltar a otro. */
var _AUD_LAYOUT = {
  santander:   { fecha: 0, monto: 2, neg: false, obs: 4 },  // [2]=Retiros, siempre positivo
  amex:        { fecha: 0, monto: 1, neg: false, obs: 3 },  // [1]=Movimientos, positivo=cargo
  mercadopago: { fecha: 1, monto: 2, neg: true,  obs: 8 }   // [2]=Cobro, negativo=egreso
};

function _auditFechaISO(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, 'America/Mexico_City', 'yyyy-MM-dd');
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  // Caja Chica mezcla Date, 'dd/mm/yyyy' y texto libre ('REMANENTE 2025') en la
  // MISMA columna. Lo que no sea fecha se descarta: no se inventa.
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}
function _auditDias(a, b) {   // diferencia en días entre dos 'yyyy-MM-dd'
  var da = new Date(a + 'T12:00:00'), db = new Date(b + 'T12:00:00');
  if (isNaN(da) || isNaN(db)) return 9999;
  return Math.abs(Math.round((da - db) / 86400000));
}
function _auditNum(v) {
  var n = (typeof v === 'number') ? v : parseFloat(String(v || '').replace(/[$,\s]/g, ''));
  return isFinite(n) ? Math.round(Math.abs(n) * 100) / 100 : 0;   // signo fuera + ruido flotante fuera
}

/* Índice de los movimientos de SALIDA que existen HOY en cada destino.
   Devuelve { santander:[{fecha,monto,obs}], amex:[…], mercadopago:[…], cajachica:[…] }
   Solo getValues(): ni un setValue en todo el archivo. */
function _auditIndiceBanco() {
  var out = { santander: [], amex: [], mercadopago: [], cajachica: [] };
  Object.keys(_AUD_LAYOUT).forEach(function (key) {
    try {
      var sh = _bankSheetByKey(key); if (!sh) return;
      var lr = sh.getLastRow(); if (lr < 2) return;
      var L = _AUD_LAYOUT[key];
      var vals = sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues();
      for (var i = 0; i < vals.length; i++) {
        var raw = vals[i][L.monto];
        var n = (typeof raw === 'number') ? raw : parseFloat(String(raw || '').replace(/[$,\s]/g, ''));
        if (!isFinite(n) || n === 0) continue;
        // Solo salidas: en MP el egreso es negativo; en Santander/AMEX, positivo
        // (un Movimiento AMEX negativo es un pago a la tarjeta, no un gasto).
        if (L.neg ? !(n < 0) : !(n > 0)) continue;
        var f = _auditFechaISO(vals[i][L.fecha]); if (!f) continue;
        out[key].push({ fecha: f, monto: Math.round(Math.abs(n) * 100) / 100, obs: String(vals[i][L.obs] || '') });
      }
    } catch (e) { Logger.log('_auditIndiceBanco ' + key + ': ' + e.message); }
  });
  try {
    var cc = getCajaChicaSheet();
    if (cc) {
      var d = cc.getDataRange().getValues();
      if (d.length > 1) {
        var H = d[0].map(function (h) { return String(h).trim().toUpperCase(); });
        var iF = H.indexOf('FECHA'), iC = H.indexOf('CONCEPTO'), iS = H.indexOf('SALIDA');
        for (var r = 1; r < d.length; r++) {
          if (iS < 0) break;
          var m = _auditNum(d[r][iS]); if (!m) continue;
          var f2 = iF > -1 ? _auditFechaISO(d[r][iF]) : ''; if (!f2) continue;
          out.cajachica.push({ fecha: f2, monto: m, obs: iC > -1 ? String(d[r][iC] || '') : '' });
        }
      }
    }
  } catch (e) { Logger.log('_auditIndiceBanco caja: ' + e.message); }
  return out;
}

/* La fecha del movimiento ETIQUETADO más antiguo = el día que el ERP empezó a
   etiquetar. Antes de esa fecha, que un egreso no tenga etiqueta NO prueba nada:
   su movimiento se cargó a mano y nunca la llevó. Se calcula, no se hardcodea:
   así se recalibra solo. Sin ninguna etiqueta devuelve '' → no se juzga nada. */
function _auditEraEtiquetas(idx) {
  var min = '';
  Object.keys(idx).forEach(function (k) {
    idx[k].forEach(function (m) {
      if (m.obs.indexOf('[Egreso #') < 0) return;
      if (!min || m.fecha < min) min = m.fecha;
    });
  });
  return min;
}

function _auditTieneEtiqueta(idx, egId) {
  var keys = Object.keys(idx);
  for (var k = 0; k < keys.length; k++) {
    var arr = idx[keys[k]];
    for (var i = 0; i < arr.length; i++) if (_egBankTagMatch(arr[i].obs, egId)) return true;
  }
  return false;
}

/* ¿Existe el movimiento, aunque no lleve etiqueta? Cruza por FECHA + MONTO.
   Tres cosas aprendidas de los datos reales, y sin las tres vuelve el falso positivo:
     · ±2 días: la domiciliación de AMEX cayó 1 día después de la fecha del egreso.
     · N:1: un egreso de $76,737.31 salió como DOS cargos (49,900 + 26,837.31).
       Se prueban pares; no se hace subset-sum general (explota y no hace falta).
     · abs() + redondeo a 2: los saldos traen ruido flotante (4.6e-10 en vez de 0).
   Devuelve {hit:true, como:'exacto'|'suma2', filas:[…]} o {hit:false}. */
function _auditBuscaMovimiento(filas, fecha, monto, tol) {
  tol = (tol == null) ? 2 : tol;
  var m = Math.round(Math.abs(monto) * 100) / 100;
  var cerca = filas.filter(function (x) { return _auditDias(x.fecha, fecha) <= tol; });
  for (var i = 0; i < cerca.length; i++) {
    if (Math.abs(cerca[i].monto - m) < 0.011) return { hit: true, como: 'exacto', filas: [cerca[i]] };
  }
  for (var a = 0; a < cerca.length; a++) {
    for (var b = a + 1; b < cerca.length; b++) {
      if (Math.abs((cerca[a].monto + cerca[b].monto) - m) < 0.011)
        return { hit: true, como: 'suma2', filas: [cerca[a], cerca[b]] };
    }
  }
  return { hit: false };
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

  // ── ÚLTIMO CANDADO, y el que de verdad importa ─────────────────────────────
  // Esta función RE-CREA la fila del banco. Si el movimiento YA existe, no está
  // reparando: está DUPLICANDO dinero, y _egresoResyncBanco no lo puede notar
  // (borra por etiqueta, y a un movimiento histórico sin etiqueta no lo ve).
  // Ya pasó una vez: la auditoría vieja juzgaba solo por etiqueta y proponía
  // "reparar" 628 movimientos que existían — $11.9M de duplicados.
  // Por eso NO basta con que la auditoría lo liste: se vuelve a buscar por
  // fecha+monto AQUÍ, contra el banco, justo antes de escribir.
  var idxRep = _auditIndiceBanco();
  var plan = [], rechazados = [];
  ids.forEach(function (id) {
    var h = porId[id];
    if (!h) { rechazados.push(id + ': no aparece como huérfano en ' + anio + ' (¿ya se reparó?)'); return; }
    if (!h.reparable) { rechazados.push(id + ': forma de pago "' + h.formaPago + '" fuera del ruteo. Corrige el egreso primero — no voy a adivinar el banco.'); return; }
    var arr = idxRep[h.destinoEsperado];
    if (arr) {
      var mv = _auditBuscaMovimiento(arr, h.fecha, h.monto);
      if (mv.hit) {
        rechazados.push(id + ': ALTO — en ' + h.destinoEsperado + ' YA hay un movimiento de $' + h.monto.toFixed(2) +
          ' el ' + mv.filas[0].fecha + (mv.como === 'suma2' ? ' (partido en 2 filas)' : '') +
          '. Solo le falta la etiqueta. Repararlo DUPLICARÍA el dinero. No se toca.');
        return;
      }
    }
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
