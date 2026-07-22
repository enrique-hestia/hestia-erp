/* ═══════════════════════════════════════════════════════════════════════════
 * ventas.gs — ÓRDENES DE VENTA (pipeline cotización → orden → facturada → pagada)
 * ---------------------------------------------------------------------------
 * COEXISTE con la captura directa de ingresos: una OV es OPCIONAL. Sirve para
 * ventas que se PACTAN antes de cobrar (cotización que el paciente aprueba) y es
 * el ancla contra la que Mercado Pago concilia por SKU.
 *
 * Vive en SHEET_ID (libro estable, NO particionado por año) porque una venta
 * cruza meses/años entre cotizar y pagar.
 *
 * ⚠ FOLIO OV-##### ESTABLE de por vida: NO cambia al convertir cotización→orden.
 *   Aprendimos que renumerar folios a media vida provoca "drift" y descuadres
 *   (ver reference_etiqueta_banco_no_prueba). El campo `Tipo` marca la etapa; el
 *   folio es el mismo desde que nace la cotización hasta que se paga.
 *
 * Escritura gateada con el MISMO permiso que la captura de ingresos (editar_ingresos):
 * si alguien puede capturar una venta, puede levantar/editar una orden de venta.
 * ═══════════════════════════════════════════════════════════════════════════ */

var OV_TAB   = 'Ordenes_Venta';
var OV_PERM  = 'editar_ingresos';
var OV_VER   = '1.0';
// Etapas del pipeline. 'cancelada' es terminal desde cualquier punto.
var OV_ESTADOS = ['borrador', 'enviada', 'aprobada', 'orden', 'facturada', 'pagada', 'cancelada'];

function _ovNum(v){ var n = parseFloat(String(v == null ? '' : v).replace(/[$,]/g, '')); return isNaN(n) ? 0 : n; }
function _ovPad2(n){ n = String(n); return n.length < 2 ? '0' + n : n; }
function _ovHoy(){
  var d = new Date();
  return d.getFullYear() + '-' + _ovPad2(d.getMonth() + 1) + '-' + _ovPad2(d.getDate());
}

/* Hoja + encabezados. Orden de columnas SAGRADO (posicional). Columnas nuevas al final. */
function _ovSheet(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(OV_TAB);
  if (!sh){
    sh = ss.insertSheet(OV_TAB);
    sh.appendRow(['Folio', 'Tipo', 'Estatus', 'Fecha', 'Vigencia', 'Paciente', 'Origen', 'Lista',
                  'Moneda', 'Lineas', 'Subtotal', 'Total', 'Notas', 'IngresoOP', 'CFDI',
                  'Usuario', 'Timestamp', 'Historial']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#f3f4f6');
  }
  return sh;
}
// Índices posicionales de la hoja.
var _OV = { folio:0, tipo:1, estatus:2, fecha:3, vigencia:4, paciente:5, origen:6, lista:7,
            moneda:8, lineas:9, subtotal:10, total:11, notas:12, ingresoOP:13, cfdi:14,
            usuario:15, ts:16, historial:17 };

/* Siguiente folio OV-#####. Bajo el lock del que llama (saveOrdenVenta lo toma). */
function _ovNextFolio(sh){
  var data = sh.getDataRange().getValues(), max = 0;
  for (var i = 1; i < data.length; i++){
    var m = String(data[i][0] || '').match(/OV-(\d+)/);
    if (m){ var n = parseInt(m[1], 10); if (n > max) max = n; }
  }
  return 'OV-' + String(max + 1).padStart(5, '0');
}

/* Normaliza y RECALCULA totales de las líneas server-side (no confía en el front). */
function _ovNormLineas(raw){
  var lineas = [], subtotal = 0;
  (raw || []).forEach(function(l){
    var pvp = _ovNum(l.pvp), cant = _ovNum(l.cantidad) || 1, desc = _ovNum(l.descuento);
    var sub = pvp * cant * (1 - desc / 100);
    subtotal += sub;
    lineas.push({ categoria: String(l.categoria || l.cat || '').trim(),
                  producto: String(l.producto || '').trim(),
                  sku: String(l.sku || '').trim(),
                  pvp: pvp, descuento: desc, cantidad: cant,
                  subtotal: Math.round(sub * 100) / 100 });
  });
  return { lineas: lineas, subtotal: Math.round(subtotal * 100) / 100 };
}

/* Localiza la fila (1-based) de un folio. 0 si no existe. */
function _ovFindRow(sh, folio){
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) if (String(data[i][0] || '').trim() === String(folio).trim()) return i + 1;
  return 0;
}

function _ovParseFila(r){
  var lineas = []; try { lineas = r[_OV.lineas] ? JSON.parse(r[_OV.lineas]) : []; } catch(e){ lineas = []; }
  var hist = []; try { hist = r[_OV.historial] ? JSON.parse(r[_OV.historial]) : []; } catch(e){ hist = []; }
  return { folio: String(r[_OV.folio] || ''), tipo: String(r[_OV.tipo] || 'cotizacion'),
           estatus: String(r[_OV.estatus] || 'borrador'), fecha: String(r[_OV.fecha] || ''),
           vigencia: String(r[_OV.vigencia] || ''), paciente: String(r[_OV.paciente] || ''),
           origen: String(r[_OV.origen] || ''), lista: String(r[_OV.lista] || ''),
           moneda: String(r[_OV.moneda] || 'MX'), lineas: lineas,
           subtotal: _ovNum(r[_OV.subtotal]), total: _ovNum(r[_OV.total]),
           notas: String(r[_OV.notas] || ''), ingresoOP: String(r[_OV.ingresoOP] || ''),
           cfdi: String(r[_OV.cfdi] || ''), usuario: String(r[_OV.usuario] || ''),
           timestamp: r[_OV.ts], historial: hist };
}

/* ─────────────────────────── LECTURA (GET) ─────────────────────────── */
/* readOrdenesVenta({estatus, tipo, paciente, folio}) — solo lectura. */
function readOrdenesVenta(params){
  try{
    params = params || {};
    var sh = _ovSheet(), data = sh.getDataRange().getValues();
    var fEst = String(params.estatus || '').trim().toLowerCase();
    var fTipo = String(params.tipo || '').trim().toLowerCase();
    var fPac = String(params.paciente || '').trim().toLowerCase();
    var fFolio = String(params.folio || '').trim().toLowerCase();
    var rows = [];
    for (var i = 1; i < data.length; i++){
      var o = _ovParseFila(data[i]);
      if (!o.folio) continue;
      if (fFolio && o.folio.toLowerCase() !== fFolio) continue;
      if (fEst && o.estatus.toLowerCase() !== fEst) continue;
      if (fTipo && o.tipo.toLowerCase() !== fTipo) continue;
      if (fPac && o.paciente.toLowerCase().indexOf(fPac) < 0) continue;
      rows.push(o);
    }
    rows.sort(function(a, b){ return String(b.folio).localeCompare(String(a.folio)); });
    return { ok: true, version: OV_VER, estados: OV_ESTADOS, rows: rows };
  }catch(ex){ return { ok: false, error: ex.message }; }
}

/* ─────────────────────────── ALTA / EDICIÓN (POST) ─────────────────────────── */
/* saveOrdenVenta(body) — crea (sin folio) o actualiza (con folio). Anti-blanqueo:
   al actualizar NO pisa IngresoOP/CFDI/Historial (se manejan por sus propios flujos). */
function saveOrdenVenta(body){
  try{
    if (!_tokenHasPermission(body.token || '', OV_PERM))
      return { ok: false, error: 'Sin permiso para gestionar ventas (' + OV_PERM + ').' };
    var norm = _ovNormLineas(body.lineas);
    if (!norm.lineas.length) return { ok: false, error: 'La orden no tiene productos.' };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return { ok: false, error: 'Otro guardado está en curso. Reintenta.' };
    try{
      var sh = _ovSheet();
      var total = norm.subtotal;   // hoy total == subtotal; el impuesto/CFDI se refleja en E4.
      var usuario = String(body.usuario || '').trim();
      var folio = String(body.folio || '').trim();

      if (!folio){
        // ── ALTA ──
        folio = _ovNextFolio(sh);
        var estatus = OV_ESTADOS.indexOf(String(body.estatus || '')) > -1 ? body.estatus : 'borrador';
        var hist = [{ estatus: estatus, usuario: usuario, ts: new Date().toISOString() }];
        sh.appendRow([folio, 'cotizacion', estatus, body.fecha || _ovHoy(), body.vigencia || '',
                      body.paciente || '', body.origen || '', body.lista || '', body.moneda || 'MX',
                      JSON.stringify(norm.lineas), norm.subtotal, total, body.notas || '',
                      '', '', usuario, new Date(), JSON.stringify(hist)]);
        try{ logAudit(usuario, 'Ventas', 'Crear cotización', folio, '', '', body.paciente || ''); }catch(e){}
        return { ok: true, folio: folio, tipo: 'cotizacion', estatus: estatus, total: total };
      }

      // ── EDICIÓN ──
      var rn = _ovFindRow(sh, folio);
      if (!rn) return { ok: false, error: 'No existe la orden ' + folio + '.' };
      var prev = _ovParseFila(sh.getRange(rn, 1, 1, 18).getValues()[0]);
      if (prev.estatus === 'cancelada') return { ok: false, error: 'La orden ' + folio + ' está cancelada; no se edita.' };
      if (prev.ingresoOP) return { ok: false, error: 'La orden ' + folio + ' ya generó el ingreso ' + prev.ingresoOP + '; para cambiarla, edita el ingreso.' };
      // Escritura celda a celda de SOLO los campos editables (no toca IngresoOP/CFDI/Historial).
      sh.getRange(rn, _OV.fecha + 1).setValue(body.fecha || prev.fecha);
      sh.getRange(rn, _OV.vigencia + 1).setValue(body.vigencia != null ? body.vigencia : prev.vigencia);
      sh.getRange(rn, _OV.paciente + 1).setValue(body.paciente != null ? body.paciente : prev.paciente);
      sh.getRange(rn, _OV.origen + 1).setValue(body.origen != null ? body.origen : prev.origen);
      sh.getRange(rn, _OV.lista + 1).setValue(body.lista != null ? body.lista : prev.lista);
      sh.getRange(rn, _OV.moneda + 1).setValue(body.moneda || prev.moneda);
      sh.getRange(rn, _OV.lineas + 1).setValue(JSON.stringify(norm.lineas));
      sh.getRange(rn, _OV.subtotal + 1).setValue(norm.subtotal);
      sh.getRange(rn, _OV.total + 1).setValue(total);
      sh.getRange(rn, _OV.notas + 1).setValue(body.notas != null ? body.notas : prev.notas);
      sh.getRange(rn, _OV.usuario + 1).setValue(usuario || prev.usuario);
      sh.getRange(rn, _OV.ts + 1).setValue(new Date());
      try{ logAudit(usuario, 'Ventas', 'Editar orden', folio, 'total', String(prev.total), String(total)); }catch(e){}
      return { ok: true, folio: folio, tipo: prev.tipo, estatus: prev.estatus, total: total };
    } finally { try{ lock.releaseLock(); }catch(e){} }
  }catch(ex){ return { ok: false, error: ex.message }; }
}

/* ─────────────────────────── CAMBIO DE ESTATUS (POST) ─────────────────────────── */
/* cambiarEstatusOV({folio, estatus, usuario, token}) — transición controlada + historial.
   Al pasar a 'orden' se marca Tipo='orden' (convertir cotización→orden, MISMO folio). */
function cambiarEstatusOV(body){
  try{
    if (!_tokenHasPermission(body.token || '', OV_PERM))
      return { ok: false, error: 'Sin permiso para gestionar ventas (' + OV_PERM + ').' };
    var nuevo = String(body.estatus || '').trim().toLowerCase();
    if (OV_ESTADOS.indexOf(nuevo) < 0) return { ok: false, error: 'Estatus inválido: ' + nuevo + '.' };
    var sh = _ovSheet(), rn = _ovFindRow(sh, body.folio);
    if (!rn) return { ok: false, error: 'No existe la orden ' + body.folio + '.' };
    var prev = _ovParseFila(sh.getRange(rn, 1, 1, 18).getValues()[0]);
    if (prev.estatus === 'cancelada') return { ok: false, error: 'La orden está cancelada.' };
    var hist = prev.historial || [];
    hist.push({ estatus: nuevo, de: prev.estatus, usuario: String(body.usuario || ''), ts: new Date().toISOString() });
    sh.getRange(rn, _OV.estatus + 1).setValue(nuevo);
    if (nuevo === 'orden' || nuevo === 'facturada' || nuevo === 'pagada') sh.getRange(rn, _OV.tipo + 1).setValue('orden');
    sh.getRange(rn, _OV.historial + 1).setValue(JSON.stringify(hist));
    sh.getRange(rn, _OV.ts + 1).setValue(new Date());
    try{ logAudit(String(body.usuario || ''), 'Ventas', 'Cambiar estatus', body.folio, 'estatus', prev.estatus, nuevo); }catch(e){}
    return { ok: true, folio: body.folio, estatus: nuevo };
  }catch(ex){ return { ok: false, error: ex.message }; }
}

/* ─────────────────────────── CANCELAR (POST) ─────────────────────────── */
function cancelarOrdenVenta(body){
  try{
    if (!_tokenHasPermission(body.token || '', OV_PERM))
      return { ok: false, error: 'Sin permiso para gestionar ventas (' + OV_PERM + ').' };
    var sh = _ovSheet(), rn = _ovFindRow(sh, body.folio);
    if (!rn) return { ok: false, error: 'No existe la orden ' + body.folio + '.' };
    var prev = _ovParseFila(sh.getRange(rn, 1, 1, 18).getValues()[0]);
    if (prev.ingresoOP) return { ok: false, error: 'La orden ya generó el ingreso ' + prev.ingresoOP + '. Cancela el ingreso, no la orden.' };
    var hist = prev.historial || [];
    hist.push({ estatus: 'cancelada', de: prev.estatus, usuario: String(body.usuario || ''), motivo: String(body.motivo || ''), ts: new Date().toISOString() });
    sh.getRange(rn, _OV.estatus + 1).setValue('cancelada');
    sh.getRange(rn, _OV.historial + 1).setValue(JSON.stringify(hist));
    try{ logAudit(String(body.usuario || ''), 'Ventas', 'Cancelar orden', body.folio, '', prev.estatus, 'cancelada · ' + (body.motivo || '')); }catch(e){}
    return { ok: true, folio: body.folio, estatus: 'cancelada' };
  }catch(ex){ return { ok: false, error: ex.message }; }
}

/* ─────────────────────────── SETUP: ítem de menú ───────────────────────────
 * Inserta "Ventas / Órdenes" (id=ordenes-venta) bajo Finanzas si no existe. Seguro
 * de correr varias veces (idempotente: no duplica). Correr UNA vez tras desplegar. */
function setupMenuVentas(){
  try{
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('Menu');
    if (!sh) return { ok:false, error:'No existe la hoja Menu.' };
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) if (String(vals[i][0] || '').trim() === 'ordenes-venta') return { ok:true, yaExiste:true };
    var ncol = vals[0].length, row = new Array(ncol);
    row[0] = 'ordenes-venta'; row[1] = 'finanzas'; row[2] = 'Ventas / Órdenes';
    row[3] = 'finanzas'; row[4] = 'file-text'; row[5] = 115; row[6] = 'vista';
    row[7] = 'OrdenVenta'; if (ncol > 8) row[8] = true;
    sh.appendRow(row);
    return { ok:true, creado:true };
  }catch(ex){ return { ok:false, error:ex.message }; }
}

/* ─────────────────────────── GENERAR INGRESO DESDE OV (POST) — E3 ───────────────
 * Puente: construye el payload de saveIngreso a partir de las líneas de la OV y lo
 * llama (reutiliza TODA la lógica de dinero: folio OP, bancos, comisiones, CxC). NO
 * duplica nada. Al éxito liga IngresoOP y mueve la OV a 'facturada'.
 * Idempotente: si la OV ya tiene IngresoOP, se niega (no genera un segundo ingreso). */
function generarIngresoDesdeOV(body){
  try{
    if (!_tokenHasPermission(body.token || '', OV_PERM))
      return { ok: false, error: 'Sin permiso para generar el ingreso (' + OV_PERM + ').' };
    if (typeof saveIngreso !== 'function') return { ok: false, error: 'Falta desplegar finance.gs (saveIngreso).' };
    var sh = _ovSheet(), rn = _ovFindRow(sh, body.folio);
    if (!rn) return { ok: false, error: 'No existe la orden ' + body.folio + '.' };
    var ov = _ovParseFila(sh.getRange(rn, 1, 1, 18).getValues()[0]);
    if (ov.estatus === 'cancelada') return { ok: false, error: 'La orden está cancelada.' };
    if (ov.ingresoOP) return { ok: false, error: 'La orden ' + body.folio + ' ya generó el ingreso ' + ov.ingresoOP + '.' };
    if (!ov.lineas.length) return { ok: false, error: 'La orden no tiene productos.' };

    var payload = {
      token: body.token, usuario: body.usuario,
      paciente: ov.paciente, fecha: body.fecha || _ovHoy(),
      formaPago: body.formaPago || '', moneda: ov.moneda,
      origenExterno: ov.origen,
      observaciones: 'Generado desde ' + ov.folio + (body.observaciones ? ' · ' + body.observaciones : ''),
      pagos: body.pagos, pagado: body.pagado,
      factura: body.factura || '', poliza: body.poliza || '',
      razonSocial: body.razonSocial || '',
      lineas: ov.lineas.map(function(l){
        return { categoria: l.categoria, producto: l.producto, sku: l.sku,
                 pvp: l.pvp, descuento: l.descuento, cantidad: l.cantidad };
      })
    };
    var res = saveIngreso(payload);
    if (!res || !res.ok) return { ok: false, error: (res && res.error) || 'No se pudo generar el ingreso.' };

    // Liga el ingreso y avanza el pipeline. El estatus final depende de si quedó saldo:
    // sin saldo → 'pagada'; con saldo (CxC generada) → 'facturada'.
    var estFinal = (res.saldoGenerado && res.saldoGenerado > 0.01) ? 'facturada' : 'pagada';
    var hist = ov.historial || [];
    hist.push({ estatus: estFinal, de: ov.estatus, usuario: String(body.usuario || ''), ingresoOP: res.op, ts: new Date().toISOString() });
    sh.getRange(rn, _OV.ingresoOP + 1).setValue(res.op);
    sh.getRange(rn, _OV.tipo + 1).setValue('orden');
    sh.getRange(rn, _OV.estatus + 1).setValue(estFinal);
    sh.getRange(rn, _OV.historial + 1).setValue(JSON.stringify(hist));
    sh.getRange(rn, _OV.ts + 1).setValue(new Date());
    try{ logAudit(String(body.usuario || ''), 'Ventas', 'Generar ingreso desde OV', body.folio, 'ingresoOP', '', res.op); }catch(e){}
    return { ok: true, folio: body.folio, op: res.op, estatus: estFinal,
             total: res.total, pagado: res.pagado, saldoGenerado: res.saldoGenerado };
  }catch(ex){ return { ok: false, error: ex.message }; }
}
