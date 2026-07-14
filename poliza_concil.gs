/**
 * poliza_concil.gs — Conciliar pólizas con ContaDigital
 *
 * Concilia los movimientos del ERP (Egresos / Ingresos) contra un EXPORT de pólizas
 * de ContaDigital (el .xlsx "crVisor") y asigna automáticamente el número de póliza
 * (Folio) al movimiento correspondiente.
 *
 * Arquitectura:
 *   - El PARSEO del .xlsx se hace en el FRONTEND (SheetJS, _loadSheetJS) y también la
 *     LÓGICA de matching (factura#/monto/fecha/nombre). Este módulo del backend solo:
 *       1) polizaConcilData(body)   → provee los movimientos del ERP del mes (fuente única,
 *                                      SIN el slice de 500 filas de readIngresosData).
 *       2) asignarPolizasLote(body) → escribe el Folio en la columna Póliza (batch).
 *       3) asignarPolizaEgreso / asignarPolizaIngreso → escritores individuales.
 *
 * Reglas de match (confirmadas con el usuario) — implementadas en el frontend:
 *   - Clave = SOLO el NÚMERO de Factura (se quita serie/letras/ceros), reforzada por
 *     Total (monto, ±tolerancia) + Fecha (±días) + nombre del Concepto.
 *   - PROV EGR → egresos ; PROV ING → ingresos.
 *   - Match casi idéntico (factura# + monto) → auto-asigna el Folio como Póliza.
 *   - Diferencias/ambigüedad → NO auto-asigna; queda para confirmación manual.
 *
 * Columna Póliza (year-aware):  Egresos → header "Póliza"  ·  BD_Ingresos → header "Poliza".
 * Permisos: escritura de egresos gated por 'editar_egresos', ingresos por 'editar_ingresos'.
 *
 * Requiere: finance.gs (readEgresosData, _egIdDeAnio/_egTabDeAnio, _ingIdDeAnio,
 *           BD_INGRESOS_TAB, _tokenHasPermission, verifyToken, logAudit).
 */

/* ── Normaliza el número de factura/folio igual que facturacion.gs (_facNormFolio):
   quita todo lo no alfanumérico, quita la serie (letras) y ceros a la izquierda,
   devuelve solo el número. La serie del emisor no discrimina. ── */
function _pcnNormFolio(s) {
  var t = String(s == null ? '' : s).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  var m = t.match(/^([A-Z]*)0*(\d+)$/);
  if (m) return m[2];
  return t;
}

// Encuentra la columna (1-based) cuyo header coincide con alguno de `names`
// (sin acentos, case-insensitive). Devuelve -1 si no existe.
function _pcnColByHeader(sheet, names) {
  var last = sheet.getLastColumn();
  if (last < 1) return -1;
  var hdrs = sheet.getRange(1, 1, 1, last).getValues()[0];
  function norm(x) {
    return String(x || '').trim().toLowerCase()
      .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
      .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u');
  }
  var wanted = [];
  for (var n = 0; n < names.length; n++) wanted.push(norm(names[n]));
  for (var c = 0; c < hdrs.length; c++) {
    var h = norm(hdrs[c]);
    for (var w = 0; w < wanted.length; w++) {
      if (h === wanted[w] || (wanted[w] && h.indexOf(wanted[w]) > -1)) return c + 1;
    }
  }
  return -1;
}

/* ── 1) Proveedor de datos: movimientos del ERP del mes ──────────────────────
   body: { anio, mes }  (mes = '' o 1..12 ; '' = todo el año)
   Devuelve egresos e ingresos con lo mínimo para el matching + un LOCATOR estable
   (egresos: rowNum ; ingresos: opId) para poder escribir la póliza después. ── */
function polizaConcilData(body) {
  try {
    var anio = parseInt((body && body.anio) || new Date().getFullYear(), 10);
    var mesNum = parseInt((body && body.mes) || 0, 10); // 0/NaN = todo el año
    if (isNaN(mesNum)) mesNum = 0;
    var mm = mesNum > 0 ? ('0' + mesNum).slice(-2) : '';

    function mesDeFecha(f) { return (f && f.length >= 7) ? f.substring(5, 7) : ''; }

    // ── Egresos: reutiliza readEgresosData (trae TODAS las filas, con _rowNum) ──
    var egOut = [], egTotal = 0;
    try {
      var eg = readEgresosData(anio);
      if (eg && eg.ok && eg.rows) {
        for (var i = 0; i < eg.rows.length; i++) {
          var e = eg.rows[i];
          if (mm && mesDeFecha(e.fecha) !== mm) continue;
          if (e.esAuto) continue; // comisiones MP automáticas: no se concilian con póliza
          egOut.push({
            loc: 'eg',
            rowNum: e._rowNum,
            id: e.id,
            fecha: e.fecha,
            monto: e.monto,
            poliza: e.poliza,
            nombre: e.proveedor || '',
            concepto: e.concepto || '',
            factura: '' // egresos no capturan número de factura → match por monto+fecha+nombre
          });
          egTotal += e.monto;
        }
      }
    } catch (eEg) { /* seguimos con lo que haya */ }

    // ── Ingresos: barrido COMPLETO de BD_Ingresos, agregado por OP (sin slice de 500) ──
    var ingOut = [], ingTotal = 0;
    try {
      var ss = SpreadsheetApp.openById(_ingIdDeAnio(anio));
      var sheet = null, all = ss.getSheets();
      for (var s = 0; s < all.length; s++) {
        if (all[s].getName() === BD_INGRESOS_TAB) { sheet = all[s]; break; }
      }
      if (sheet) {
        var raw = sheet.getDataRange().getValues();
        function num(v) { if (typeof v === 'number') return v; var n = parseFloat(String(v || '').replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
        function dt(v) { if (!v) return ''; if (v instanceof Date) return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2); return String(v); }
        // OP(0),Linea(1),Fecha(2),Paciente(3),...,TotalPagar(9),...,MontoFactMes(11),...,Factura(17),Poliza(18)
        var byOp = {};
        for (var r = 1; r < raw.length; r++) {
          var op = String(raw[r][0] || '').trim();
          if (!op) continue;
          var fecha = dt(raw[r][2]);
          if (mm && mesDeFecha(fecha) !== mm) continue;
          if (!byOp[op]) {
            byOp[op] = { loc: 'ing', opId: op, fecha: fecha,
              nombre: String(raw[r][3] || ''), factura: '', poliza: '',
              monto: 0, montoFact: 0 };
          }
          var rec = byOp[op];
          if (!rec.fecha && fecha) rec.fecha = fecha;
          rec.monto += num(raw[r][9]);
          rec.montoFact += num(raw[r][11]);
          var fac = String(raw[r][17] || '').trim();
          if (fac && !rec.factura) rec.factura = fac;
          var pol = String(raw[r][18] || '').trim();
          if (pol && !rec.poliza) rec.poliza = pol;
        }
        for (var k in byOp) { ingOut.push(byOp[k]); ingTotal += byOp[k].monto; }
      }
    } catch (eIn) { /* seguimos con lo que haya */ }

    return { ok: true, anio: anio, mes: mesNum || 0,
      egresos: egOut, ingresos: ingOut,
      totalEgresos: egTotal, totalIngresos: ingTotal };
  } catch (ex) {
    return { ok: false, error: ex.message, egresos: [], ingresos: [] };
  }
}

/* ── 2) Escritor individual: póliza de un EGRESO ─────────────────────────────
   body: { token, anio, rowNum, egId, poliza, usuario }
   Verifica que la fila (rowNum) siga siendo el egreso esperado (col A = egId) antes
   de escribir, para no clavar la póliza en la fila equivocada si la hoja cambió. ── */
function asignarPolizaEgreso(body) {
  try {
    if (!_tokenHasPermission((body && body.token) || '', 'editar_egresos'))
      return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var anio = parseInt((body && body.anio) || new Date().getFullYear(), 10);
    var poliza = String((body && body.poliza) || '').trim();
    var rowNum = parseInt((body && body.rowNum) || 0, 10);
    if (!poliza) return { ok: false, error: 'Sin póliza.' };
    if (!rowNum || rowNum < 2) return { ok: false, error: 'Fila inválida.' };

    var ss = SpreadsheetApp.openById(_egIdDeAnio(anio));
    var tab = _egTabDeAnio(anio), sheet = null, sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) { if (sheets[i].getName() === tab) { sheet = sheets[i]; break; } }
    if (!sheet) sheet = sheets[0];
    if (!sheet) return { ok: false, error: 'Pestaña de egresos no encontrada.' };

    // Verificación de identidad (col A = ID del egreso)
    if (body && body.egId != null && String(body.egId).trim() !== '') {
      var idCel = String(sheet.getRange(rowNum, 1).getValue() || '').trim();
      if (idCel !== String(body.egId).trim())
        return { ok: false, error: 'La fila cambió (ID no coincide). Vuelve a conciliar.' };
    }

    var colPol = _pcnColByHeader(sheet, ['póliza', 'poliza']);
    if (colPol < 1) return { ok: false, error: 'Columna Póliza no encontrada en egresos.' };
    sheet.getRange(rowNum, colPol).setValue(poliza);

    try { logAudit((body && body.usuario) || 'sistema', 'Egresos', 'Asignar póliza (ContaDigital)', 'Fila ' + rowNum, '', '', poliza); } catch (ae) {}
    return { ok: true, rowNum: rowNum, poliza: poliza };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── 2b) Escritor individual: póliza de un INGRESO (por OP) ──────────────────
   body: { token, anio, opId, poliza, usuario }
   Escribe la póliza en TODAS las líneas del OP (consistente con updateIngresoFiscal). ── */
function asignarPolizaIngreso(body) {
  try {
    if (!_tokenHasPermission((body && body.token) || '', 'editar_ingresos'))
      return { ok: false, error: 'Sin autorización (editar_ingresos).' };
    var anio = parseInt((body && body.anio) || new Date().getFullYear(), 10);
    var opId = String((body && body.opId) || '').trim();
    var poliza = String((body && body.poliza) || '').trim();
    if (!opId) return { ok: false, error: 'Sin opId.' };
    if (!poliza) return { ok: false, error: 'Sin póliza.' };

    var ss = SpreadsheetApp.openById(_ingIdDeAnio(anio));
    var sheet = null, sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) { if (sheets[i].getName() === BD_INGRESOS_TAB) { sheet = sheets[i]; break; } }
    if (!sheet) return { ok: false, error: 'BD_Ingresos no encontrada.' };

    var colPol = _pcnColByHeader(sheet, ['poliza', 'póliza']);
    if (colPol < 1) return { ok: false, error: 'Columna Poliza no encontrada en ingresos.' };

    var data = sheet.getDataRange().getValues();
    var updated = 0;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0] || '').trim() === opId) {
        sheet.getRange(r + 1, colPol).setValue(poliza);
        updated++;
      }
    }
    if (!updated) return { ok: false, error: 'OP no encontrado en BD_Ingresos: ' + opId };
    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (ce) {}
    try { logAudit((body && body.usuario) || 'sistema', 'Ingresos', 'Asignar póliza (ContaDigital)', opId, '', '', poliza); } catch (ae) {}
    return { ok: true, opId: opId, poliza: poliza, updated: updated };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ── 3) Escritor BATCH: aplica todas las auto-asignaciones en una sola llamada ──
   body: {
     token, anio, usuario,
     asignaciones: [ { side:'egreso', rowNum, egId, poliza } | { side:'ingreso', opId, poliza } ]
   }
   Gatea permisos POR LADO (solo exige editar_egresos si hay egresos, etc.). Abre cada
   libro una sola vez. Devuelve conteos + resultado por item. ── */
function asignarPolizasLote(body) {
  try {
    var anio = parseInt((body && body.anio) || new Date().getFullYear(), 10);
    var token = (body && body.token) || '';
    var usuario = (body && body.usuario) || 'sistema';
    var items = (body && body.asignaciones) || [];
    if (!items.length) return { ok: false, error: 'Sin asignaciones.' };

    var hayEg = false, hayIng = false;
    for (var t = 0; t < items.length; t++) {
      if (items[t] && items[t].side === 'egreso') hayEg = true;
      if (items[t] && items[t].side === 'ingreso') hayIng = true;
    }
    var puedeEg = hayEg ? _tokenHasPermission(token, 'editar_egresos') : true;
    var puedeIng = hayIng ? _tokenHasPermission(token, 'editar_ingresos') : true;
    if (hayEg && !puedeEg) return { ok: false, error: 'Sin autorización para egresos (editar_egresos).' };
    if (hayIng && !puedeIng) return { ok: false, error: 'Sin autorización para ingresos (editar_ingresos).' };

    var resultados = [], okEg = 0, okIng = 0, fail = 0;

    // ── Egresos (un solo open de libro/pestaña) ──
    if (hayEg) {
      var ssE = SpreadsheetApp.openById(_egIdDeAnio(anio));
      var tabE = _egTabDeAnio(anio), shE = null, shsE = ssE.getSheets();
      for (var e0 = 0; e0 < shsE.length; e0++) { if (shsE[e0].getName() === tabE) { shE = shsE[e0]; break; } }
      if (!shE) shE = shsE[0];
      var colPolE = shE ? _pcnColByHeader(shE, ['póliza', 'poliza']) : -1;
      for (var ei = 0; ei < items.length; ei++) {
        var it = items[ei];
        if (!it || it.side !== 'egreso') continue;
        var rn = parseInt(it.rowNum || 0, 10);
        var pol = String(it.poliza || '').trim();
        if (!shE || colPolE < 1) { resultados.push({ side: 'egreso', rowNum: rn, ok: false, error: 'Sin hoja/columna.' }); fail++; continue; }
        if (!rn || rn < 2 || !pol) { resultados.push({ side: 'egreso', rowNum: rn, ok: false, error: 'Datos inválidos.' }); fail++; continue; }
        if (it.egId != null && String(it.egId).trim() !== '') {
          var idc = String(shE.getRange(rn, 1).getValue() || '').trim();
          if (idc !== String(it.egId).trim()) { resultados.push({ side: 'egreso', rowNum: rn, ok: false, error: 'Fila cambió.' }); fail++; continue; }
        }
        shE.getRange(rn, colPolE).setValue(pol);
        resultados.push({ side: 'egreso', rowNum: rn, ok: true, poliza: pol }); okEg++;
      }
    }

    // ── Ingresos (un solo open; escribe todas las líneas de cada OP) ──
    if (hayIng) {
      var ssI = SpreadsheetApp.openById(_ingIdDeAnio(anio));
      var shI = null, shsI = ssI.getSheets();
      for (var i0 = 0; i0 < shsI.length; i0++) { if (shsI[i0].getName() === BD_INGRESOS_TAB) { shI = shsI[i0]; break; } }
      var colPolI = shI ? _pcnColByHeader(shI, ['poliza', 'póliza']) : -1;
      var dataI = shI ? shI.getDataRange().getValues() : [];
      // índice OP -> filas (1-based) para no recorrer la hoja por cada item
      var opRows = {};
      for (var di = 1; di < dataI.length; di++) {
        var opv = String(dataI[di][0] || '').trim();
        if (!opv) continue;
        if (!opRows[opv]) opRows[opv] = [];
        opRows[opv].push(di + 1);
      }
      for (var ii = 0; ii < items.length; ii++) {
        var itI = items[ii];
        if (!itI || itI.side !== 'ingreso') continue;
        var opId = String(itI.opId || '').trim();
        var polI = String(itI.poliza || '').trim();
        if (!shI || colPolI < 1) { resultados.push({ side: 'ingreso', opId: opId, ok: false, error: 'Sin hoja/columna.' }); fail++; continue; }
        if (!opId || !polI || !opRows[opId]) { resultados.push({ side: 'ingreso', opId: opId, ok: false, error: 'OP no encontrado.' }); fail++; continue; }
        var rows = opRows[opId];
        for (var rr = 0; rr < rows.length; rr++) shI.getRange(rows[rr], colPolI).setValue(polI);
        resultados.push({ side: 'ingreso', opId: opId, ok: true, poliza: polI, lineas: rows.length }); okIng++;
      }
      try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch (ce) {}
    }

    try { logAudit(usuario, 'Conciliación', 'Asignar pólizas (ContaDigital) lote', '', '', '', 'eg=' + okEg + ' ing=' + okIng + ' fail=' + fail); } catch (ae) {}
    return { ok: true, okEgresos: okEg, okIngresos: okIng, fallidos: fail, resultados: resultados };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
