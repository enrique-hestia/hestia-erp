/* ══════════════════════════════════════════════════════════════
   prov_defaults.gs — Clasificación automática de gastos por proveedor
   ------------------------------------------------------------
   Objetivo: al capturar un egreso, si el proveedor es "frecuente"
   (o tiene una clasificación fija guardada), recomendar en automático
   Contable · Tipo · Subtipo · Concepto para evitar errores de captura
   que descuadran el Estado de Resultados.

   Dos fuentes de recomendación:
     1. FIJA  — columnas "Def Contable/Tipo/Subtipo/Concepto" en la hoja
        Proveedores (CXP_SS_ID). Si están llenas, mandan siempre (pct 100).
     2. HISTÓRICA — combinación más frecuente en BD_Egresos (año actual +
        anterior). Se sugiere con su frecuencia (veces / % de sus gastos).

   Requiere: providers.gs (CXP_SS_ID, PROV_TAB, readProveedores),
   finance.gs (readEgresosData), core.gs/finance.gs para el wiring.
   ══════════════════════════════════════════════════════════════ */

var PROV_DEF_COLS = ['Def Contable','Def Tipo','Def Subtipo','Def Concepto'];

function _pdNorm(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,' '); }

/* Localiza (o crea) las columnas de default en la hoja Proveedores.
   Devuelve {sh, idxNombre, idx:{contable,tipo,subtipo,concepto}, headers}. */
function _pdEnsureCols(create){
  var ss = SpreadsheetApp.openById(CXP_SS_ID);
  var sh = ss.getSheetByName(PROV_TAB);
  if (!sh) { if (create) { setupProveedores(); sh = ss.getSheetByName(PROV_TAB); } else return null; }
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  function find(re){ for (var i=0;i<headers.length;i++){ if(re.test(headers[i])) return i; } return -1; }
  var idxNombre = find(/nombre|raz[oó]n/i); if (idxNombre<0) idxNombre = 1;
  var idx = {
    contable: find(/def.*contable/i),
    tipo:     find(/def.*tipo/i),
    subtipo:  find(/def.*subtipo/i),
    concepto: find(/def.*concepto/i)
  };
  // Crear las que falten al final
  if (create){
    var toAdd = [];
    PROV_DEF_COLS.forEach(function(lbl){
      var kmap = {'Def Contable':'contable','Def Tipo':'tipo','Def Subtipo':'subtipo','Def Concepto':'concepto'};
      var k = kmap[lbl];
      if (idx[k] < 0){ toAdd.push(lbl); }
    });
    if (toAdd.length){
      sh.getRange(1, lastCol+1, 1, toAdd.length).setValues([toAdd])
        .setFontWeight('bold').setFontColor('#ffffff').setBackground('#7a5560');
      // re-leer headers
      headers = sh.getRange(1,1,1,lastCol+toAdd.length).getValues()[0].map(function(h){ return String(h||'').trim(); });
      idx = {
        contable: (function(){for(var i=0;i<headers.length;i++)if(/def.*contable/i.test(headers[i]))return i;return -1;})(),
        tipo:     (function(){for(var i=0;i<headers.length;i++)if(/def.*tipo/i.test(headers[i]))return i;return -1;})(),
        subtipo:  (function(){for(var i=0;i<headers.length;i++)if(/def.*subtipo/i.test(headers[i]))return i;return -1;})(),
        concepto: (function(){for(var i=0;i<headers.length;i++)if(/def.*concepto/i.test(headers[i]))return i;return -1;})()
      };
    }
  }
  return { sh:sh, idxNombre:idxNombre, idx:idx, headers:headers };
}

/* ── Lee el mapa de recomendaciones por proveedor ─────────────────
   Cachea 10 min (CacheService). Devuelve map keyed por nombre normalizado. */
function readClasificacionProveedores(force){
  try {
    var cache = CacheService.getScriptCache();
    if (!force){
      var hit = cache.get('clasifProv_v1');
      if (hit) { try { return JSON.parse(hit); } catch(e){} }
    }
    var map = {}; // nombreLower -> {nombre,contable,tipo,subtipo,concepto,fuente,veces,total,pct,combos}

    // 1) Aprendizaje histórico de BD_Egresos (año actual + anterior)
    var anioAct = new Date().getFullYear();
    [anioAct-1, anioAct].forEach(function(y){
      var eg; try { eg = readEgresosData(y); } catch(e){ return; }
      (eg.rows||[]).forEach(function(r){
        var nombre = String(r.proveedor||'').trim();
        if (!nombre) return;
        if (r.estatus === 'Cancelada') return;
        var k = _pdNorm(nombre);
        if (!map[k]) map[k] = { nombre:nombre, fuente:'', veces:0, total:0, _combos:{}, _conceptos:{} };
        var m = map[k];
        m.veces++; m.total += Number(r.monto)||0;
        var combo = String(r.contable||'')+'||'+String(r.tipo||'')+'||'+String(r.subtipo||'');
        m._combos[combo] = (m._combos[combo]||0) + 1;
        var conc = String(r.concepto||'').trim();
        if (conc) m._conceptos[conc] = (m._conceptos[conc]||0) + 1;
      });
    });

    // Resolver combo/concepto más frecuente por proveedor
    Object.keys(map).forEach(function(k){
      var m = map[k];
      var bestCombo='', bestN=0;
      Object.keys(m._combos).forEach(function(c){ if(m._combos[c]>bestN){ bestN=m._combos[c]; bestCombo=c; } });
      var parts = bestCombo.split('||');
      m.contable = parts[0]||''; m.tipo = parts[1]||''; m.subtipo = parts[2]||'';
      var bestC='', bestCN=0;
      Object.keys(m._conceptos).forEach(function(c){ if(m._conceptos[c]>bestCN){ bestCN=m._conceptos[c]; bestC=c; } });
      m.concepto = bestC;
      m.pct = m.veces>0 ? Math.round((bestN/m.veces)*100) : 0;
      m.comboVeces = bestN;
      m.fuente = 'historico';
      delete m._combos; delete m._conceptos;
    });

    // 2) Overrides FIJOS desde la hoja Proveedores
    try {
      var info = _pdEnsureCols(false);
      if (info && (info.idx.contable>-1 || info.idx.subtipo>-1)){
        var sh = info.sh;
        var lr = sh.getLastRow();
        if (lr > 1){
          var vals = sh.getRange(2,1,lr-1,sh.getLastColumn()).getValues();
          vals.forEach(function(row){
            var nombre = String(row[info.idxNombre]||'').trim();
            if (!nombre) return;
            var dc = info.idx.contable>-1 ? String(row[info.idx.contable]||'').trim() : '';
            var dt = info.idx.tipo>-1 ? String(row[info.idx.tipo]||'').trim() : '';
            var ds = info.idx.subtipo>-1 ? String(row[info.idx.subtipo]||'').trim() : '';
            var dk = info.idx.concepto>-1 ? String(row[info.idx.concepto]||'').trim() : '';
            if (!dc && !dt && !ds && !dk) return; // sin default fijo
            var k = _pdNorm(nombre);
            var prev = map[k] || { nombre:nombre, veces:0, total:0 };
            map[k] = { nombre:nombre, contable:dc, tipo:dt, subtipo:ds, concepto:dk,
                       fuente:'fijo', pct:100, veces:prev.veces||0, total:prev.total||0, comboVeces:prev.veces||0 };
          });
        }
      }
    } catch(e2){}

    var out = { ok:true, map:map, count:Object.keys(map).length, generado:new Date().toISOString() };
    try { cache.put('clasifProv_v1', JSON.stringify(out), 600); } catch(e3){}
    return out;
  } catch(ex){ return { ok:false, error:ex.message, map:{} }; }
}

/* ── Guarda una clasificación FIJA en el proveedor ────────────────
   body: {nombre, contable, tipo, subtipo, concepto, usuario} */
function guardarClasifProveedor(body){
  try {
    var nombre = String((body&&body.nombre)||'').trim();
    if (!nombre) return { ok:false, error:'Falta el nombre del proveedor.' };
    var info = _pdEnsureCols(true);
    if (!info) return { ok:false, error:'No existe la hoja Proveedores.' };
    var sh = info.sh, lr = sh.getLastRow();
    if (lr < 2) return { ok:false, error:'No hay proveedores dados de alta.' };
    var col = sh.getLastColumn();
    var vals = sh.getRange(2,1,lr-1,col).getValues();
    var targetRow = -1;
    var nk = _pdNorm(nombre);
    for (var i=0;i<vals.length;i++){
      if (_pdNorm(vals[i][info.idxNombre]) === nk){ targetRow = i+2; break; }
    }
    if (targetRow < 0) return { ok:false, error:'El proveedor "'+nombre+'" no está en el catálogo. Da de alta el proveedor primero.' };
    function setCell(idx, val){ if (idx>-1) sh.getRange(targetRow, idx+1).setValue(val==null?'':val); }
    setCell(info.idx.contable, body.contable);
    setCell(info.idx.tipo, body.tipo);
    setCell(info.idx.subtipo, body.subtipo);
    setCell(info.idx.concepto, body.concepto);
    try { CacheService.getScriptCache().remove('clasifProv_v1'); } catch(e){}
    try { logAudit((body&&body.usuario)||'sistema','Proveedores','ClasifFija',nombre,'','',
      [body.contable,body.tipo,body.subtipo,body.concepto].join(' · ')); } catch(e){}
    return { ok:true, nombre:nombre };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
