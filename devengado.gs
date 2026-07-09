/**
 * devengado.gs — Gasto DEVENGADO (informativo, no financiero)
 *
 * Proveedores de "facturación vencida" (ej. Vianto): consumes el servicio en un mes
 * pero te lo cobran los primeros días del mes siguiente. El egreso queda fechado en el
 * mes del PAGO, así que en flujo de efectivo y en el P&L cae en el mes equivocado.
 *
 * Este módulo NO toca nada financiero: ni el flujo, ni el estado de resultados, ni
 * totales. Solo recalcula, de forma INFORMATIVA, cuánto gasto CORRESPONDE a cada mes
 * (base devengado) recorriendo esos pagos al mes en que se consumió el servicio.
 *
 * Regla por proveedor (Script Property FACT_VENCIDA):
 *   { proveedor, offset (meses, típico -1), diaTope }
 *   - Si el pago cae en un día <= diaTope  -> corresponde al mes (pago + offset).
 *   - Si cae después de diaTope            -> se queda en su propio mes (fue factura del mes).
 * Override manual (gana sobre la regla): columna "Mes Devengado" (YYYY-MM) en la hoja de Egresos.
 *
 * Requiere: finance.gs (readEgresosData).
 */

// ── Config ────────────────────────────────────────────────────────────────
function _dvNorm(s){
  return String(s||'').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g,'')   // sin acentos
    .replace(/[.,]/g,'').replace(/\s+/g,' ').trim();
}

function _factVencidaCfg(){
  var raw;
  try { raw = PropertiesService.getScriptProperties().getProperty('FACT_VENCIDA'); } catch(e){ raw = null; }
  var arr = null;
  if (raw){ try { arr = JSON.parse(raw); } catch(e){ arr = null; } }
  if (!arr || !arr.length) arr = [{ proveedor:'Vianto', offset:-1, diaTope:12 }]; // semilla por defecto
  return arr;
}

function _factVencidaMap(){
  var arr = _factVencidaCfg(), m = {};
  arr.forEach(function(x){
    if (!x || !x.proveedor) return;
    m[_dvNorm(x.proveedor)] = { offset:(x.offset==null?-1:Number(x.offset)), diaTope:(x.diaTope==null?12:Number(x.diaTope)), label:String(x.proveedor) };
  });
  return m;
}

// GET: config + lista de proveedores conocidos (para el editor)
function factVencidaLeer(){
  var cfg = _factVencidaCfg();
  var provs = {};
  try {
    var y = new Date().getFullYear();
    [y, y-1].forEach(function(yy){
      try { (readEgresosData(yy).rows||[]).forEach(function(r){ if(r.proveedor) provs[r.proveedor]=1; }); } catch(e){}
    });
  } catch(e){}
  return { ok:true, config:cfg, proveedores:Object.keys(provs).sort() };
}

// POST: guardar config { config:[{proveedor,offset,diaTope}] }
function factVencidaGuardar(body){
  try {
    var arr = (body && body.config) || [];
    var limpio = [];
    (arr||[]).forEach(function(x){
      if (!x || !String(x.proveedor||'').trim()) return;
      limpio.push({ proveedor:String(x.proveedor).trim(),
        offset:(x.offset==null?-1:Number(x.offset)),
        diaTope:(x.diaTope==null?12:Number(x.diaTope)) });
    });
    PropertiesService.getScriptProperties().setProperty('FACT_VENCIDA', JSON.stringify(limpio));
    return { ok:true, config:limpio };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

// ── Helpers de mes ──────────────────────────────────────────────────────────
function _dvMesShift(ym, delta){
  var y = parseInt(ym.substring(0,4),10), m = parseInt(ym.substring(5,7),10) - 1 + (Number(delta)||0);
  y += Math.floor(m/12); m = ((m%12)+12)%12;
  return y + '-' + String(m+1).padStart(2,'0');
}

// Mes de SERVICIO de un egreso (YYYY-MM) o null si el proveedor no aplica.
function _dvMesServicio(r, map){
  var f = String(r.fecha||'').substring(0,10);
  if (f.length < 7) return null;
  var mesPago = f.substring(0,7);
  // override explícito (columna "Mes Devengado")
  var ov = String(r.mesDevengado||'').trim();
  if (/^\d{4}-\d{2}$/.test(ov)) return ov;
  var cfg = map[_dvNorm(r.proveedor)];
  if (!cfg) return null;
  var dia = parseInt(f.substring(8,10),10) || 1;
  if (cfg.diaTope && dia > cfg.diaTope) return mesPago;      // cobrado tarde -> es del propio mes
  return _dvMesShift(mesPago, cfg.offset==null?-1:cfg.offset);
}

// ── Motor informativo ────────────────────────────────────────────────────────
// Devuelve, para el rango [fi..ff], el ajuste devengado (recorrido de pagos) sin tocar totales.
//   entran  = pagos hechos FUERA del rango cuyo servicio cae DENTRO (suman al periodo)
//   salen   = pagos hechos DENTRO del rango cuyo servicio cae FUERA (restan del periodo)
//   ajuste  = entran - salen  -> correspondeAlPeriodo = egresosPagados + ajuste
//   porMes  = {'YYYY-MM':{entran,salen}} incluye recorridos internos entre meses del rango
function readGastoDevengado(fi, ff){
  try {
    fi = String(fi||'').substring(0,10); ff = String(ff||'').substring(0,10);
    if (fi.length<10 || ff.length<10){
      var hoy=new Date(); ff = hoy.getFullYear()+'-'+String(hoy.getMonth()+1).padStart(2,'0')+'-28';
      fi = ff.substring(0,4)+'-01-01';
    }
    var map = _factVencidaMap();
    if (!Object.keys(map).length) return { ok:true, aplica:false, entran:0, salen:0, ajuste:0, detalle:[], porMes:{} };

    var mesFi = fi.substring(0,7), mesFf = ff.substring(0,7);
    function enRango(ym){ return ym >= mesFi && ym <= mesFf; }
    var yFrom = parseInt(fi.substring(0,4),10), yTo = parseInt(ff.substring(0,4),10);

    var entran=0, salen=0, detalle=[], porMes={};
    function bucket(m){ if(!porMes[m]) porMes[m]={ mes:m, entran:0, salen:0 }; return porMes[m]; }

    // Dedupe de spreadsheets: un año SIN hoja propia (ej. 2027) cae de vuelta al de 2026
    // (readEgresosData: EGRESOS_IDS[anio] || EGRESOS_SS_2026). Sin esto, el egreso de Vianto
    // se leería dos veces (2026 y 2027→2026) y saldría DUPLICADO.
    var _seenSS = {};
    for (var y=yFrom-1; y<=yTo+1; y++){
      var _ssId = (typeof EGRESOS_IDS !== 'undefined' && EGRESOS_IDS[y]) ? EGRESOS_IDS[y]
                : (typeof EGRESOS_SS_2026 !== 'undefined' ? EGRESOS_SS_2026 : ('year-'+y));
      if (_seenSS[_ssId]) continue;   // ya leímos ese spreadsheet
      _seenSS[_ssId] = 1;
      var eg; try { eg = readEgresosData(y); } catch(e){ continue; }
      (eg.rows||[]).forEach(function(r){
        if (r.estatus === 'Cancelada') return;
        var ms = _dvMesServicio(r, map);
        if (!ms) return;
        var f = String(r.fecha||'').substring(0,10); if (f.length<7) return;
        var mp = f.substring(0,7);
        if (mp === ms) return;                      // no se recorre
        var monto = Number(r.monto)||0; if (!monto) return;
        var pagoIn = enRango(mp), servIn = enRango(ms);
        if (pagoIn && !servIn){ salen += monto; bucket(mp).salen += monto;
          detalle.push({ proveedor:r.proveedor, fecha:f, monto:monto, mesPago:mp, mesServicio:ms, tipo:'sale' }); }
        else if (!pagoIn && servIn){ entran += monto; bucket(ms).entran += monto;
          detalle.push({ proveedor:r.proveedor, fecha:f, monto:monto, mesPago:mp, mesServicio:ms, tipo:'entra' }); }
        else if (pagoIn && servIn){ bucket(mp).salen += monto; bucket(ms).entran += monto;
          detalle.push({ proveedor:r.proveedor, fecha:f, monto:monto, mesPago:mp, mesServicio:ms, tipo:'interno' }); }
      });
    }
    detalle.sort(function(a,b){ return a.fecha < b.fecha ? -1 : 1; });

    var hoy = new Date();
    var mesHoy = hoy.getFullYear()+'-'+String(hoy.getMonth()+1).padStart(2,'0');
    var parcial = mesFf >= mesHoy;   // el pago del mes siguiente puede no existir aún

    return { ok:true, aplica:true, entran:entran, salen:salen, ajuste:(entran-salen),
             detalle:detalle, porMes:porMes, parcial:parcial,
             proveedores:Object.keys(map).map(function(k){ return map[k].label; }) };
  } catch(ex){ return { ok:false, error:ex.message+' (L:'+(ex.lineNumber||'?')+')' }; }
}
