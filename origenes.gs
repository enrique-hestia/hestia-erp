/* ══════════════════════════════════════════════════════════════════════
   ORÍGENES EXTERNOS — atribución del "dueño" de un ingreso externo.

   Los ingresos del bucket "Externos" (CicloAltaBaja/Categoría con 'extern'
   o que matchea una AGENCIA) no tienen dueño estructurado: hoy se adivina
   por substrings. Este módulo agrega una dimensión de dueño ("Origen
   externo") con un CATÁLOGO editable (hoja Origenes_Externos) y un
   back-fill histórico asistido.

   Contrato con el frontend:
     GET  action=origenes                    → readOrigenes()
     GET  action=sugerirOrigenes&anio=YYYY    → sugerirOrigenesHistorico(anio)
     POST action=saveOrigen                   → saveOrigen(body)
     POST action=deleteOrigen                 → deleteOrigen(body)
     POST action=aplicarOrigenesHistorico     → aplicarOrigenesHistorico(body)

   Dependencias (mismo despliegue): SHEET_ID (config.gs), INGRESOS_SS_ID /
   _ingIdDeAnio / BD_INGRESOS_TAB / _tokenHasPermission / _ingColEnsure
   (finance.gs), _summaryAgencias / _summaryEsAgencia (summary.gs),
   jsonResponse (core.gs).

   IMPORTANTE: correr setupOrigenesExternos() UNA vez desde el editor de
   Apps Script para crear/sembrar la hoja. La columna OrigenExterno en
   BD_Ingresos se crea sola en la primera escritura (ver _ingColEnsure).
   ══════════════════════════════════════════════════════════════════════ */

var ORIGENES_TAB     = 'Origenes_Externos';
var ORIGENES_HEADERS = ['ID','Nombre','Tipo','Alias','Activo','Notas','CreadoEn'];

/* ── Normalización: minúsculas, sin acentos, puntuación→espacio, colapsada.
   Escrita SIN String.prototype.normalize a propósito: usa escapes \u para
   ser idéntica y ejecutable tanto en V8 como en el simulador ES3 (cscript). */
function _origNorm(s){
  s = String(s == null ? '' : s).toLowerCase();
  s = s.replace(/[áàäâã]/g, 'a')
       .replace(/[éèëê]/g, 'e')
       .replace(/[íìïî]/g, 'i')
       .replace(/[óòöôõ]/g, 'o')
       .replace(/[úùüû]/g, 'u')
       .replace(/ñ/g, 'n')
       .replace(/ç/g, 'c');
  s = s.replace(/[^a-z0-9]+/g, ' ').replace(/^\s+|\s+$/g, '');
  return s;
}

/* ── Mejor coincidencia de un texto contra el catálogo de orígenes.
   Prueba el Nombre normalizado (→ confianza 'alta') y cada Alias
   normalizado (→ 'baja') como SUBSTRING del texto normalizado.
   'alta' siempre gana sobre 'baja'. Devuelve {nombre,tipo,confianza}
   (confianza '' si nada matchea). Ignora orígenes con activo===false. */
function _origMatch(texto, origenesArr){
  var t = _origNorm(texto);
  var baja = null;
  if (!t || !origenesArr) return { nombre:'', tipo:'', confianza:'' };
  for (var i = 0; i < origenesArr.length; i++){
    var o = origenesArr[i] || {};
    if (o.activo === false) continue;
    var nom = String(o.nombre || '');
    var nn  = _origNorm(nom);
    if (nn && t.indexOf(nn) > -1){
      return { nombre:nom, tipo:String(o.tipo || ''), confianza:'alta' };
    }
    if (!baja){
      var aliases = String(o.alias || '').split(',');
      for (var a = 0; a < aliases.length; a++){
        var al = _origNorm(aliases[a]);
        if (al && t.indexOf(al) > -1){
          baja = { nombre:nom, tipo:String(o.tipo || ''), confianza:'baja' };
          break;
        }
      }
    }
  }
  return baja || { nombre:'', tipo:'', confianza:'' };
}

/* ── Helpers de ID ORIG-00001 ─────────────────────────────────────────── */
function _origPad(n){ var s = String(n); while (s.length < 5) s = '0' + s; return s; }
function _origMaxIdNum(data){
  var mx = 0;
  for (var i = 1; i < data.length; i++){
    var m = String((data[i] && data[i][0]) || '').match(/ORIG-(\d+)/);
    if (m){ var v = parseInt(m[1], 10); if (v > mx) mx = v; }
  }
  return mx;
}

/* ── Verdad tolerante para la columna Activo (celda vacía = activo). ───── */
function _origTruthy(v){
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true; // celda vacía → activo por defecto
  return !(s === 'false' || s === 'no' || s === '0' || s === 'inactivo');
}

/* ── Setup: crea/estiliza la hoja y SIEMBRA una fila por cada AGENCIA
   (Tipo='Agencia'). Idempotente: no duplica por nombre. Correr una vez. */
function setupOrigenesExternos(){
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(ORIGENES_TAB);
    var created = false;
    if (!sh){ sh = ss.insertSheet(ORIGENES_TAB); created = true; }

    sh.getRange(1, 1, 1, ORIGENES_HEADERS.length).setValues([ORIGENES_HEADERS]);
    sh.getRange(1, 1, 1, ORIGENES_HEADERS.length)
      .setFontWeight('bold').setBackground('#c46a7a').setFontColor('#ffffff');
    sh.setFrozenRows(1);

    var data = sh.getDataRange().getValues();
    var existing = {};
    for (var i = 1; i < data.length; i++){ existing[_origNorm(data[i][1])] = true; }

    var ags = (typeof _summaryAgencias === 'function') ? _summaryAgencias() : ['REPROVIDA'];
    var counter = _origMaxIdNum(data);
    var toAdd = [];
    for (var a = 0; a < ags.length; a++){
      var nom = String(ags[a] || '').trim();
      if (!nom || existing[_origNorm(nom)]) continue;
      counter++;
      toAdd.push(['ORIG-' + _origPad(counter), nom, 'Agencia', '', true, 'Sembrado desde AGENCIAS', new Date()]);
      existing[_origNorm(nom)] = true;
    }
    if (toAdd.length) sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, ORIGENES_HEADERS.length).setValues(toAdd);

    return { ok:true, created:created, seeded:toAdd.length, tab:ORIGENES_TAB };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Lectura del catálogo ─────────────────────────────────────────────── */
function readOrigenes(){
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(ORIGENES_TAB);
    if (!sh) return { ok:true, origenes:[] };
    var data = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < data.length; i++){
      var r = data[i];
      if (!String(r[0] || '').trim() && !String(r[1] || '').trim()) continue;
      out.push({
        rowNum: i + 1,
        id:     String(r[0] || ''),
        nombre: String(r[1] || ''),
        tipo:   String(r[2] || ''),
        alias:  String(r[3] || ''),
        activo: _origTruthy(r[4]),
        notas:  String(r[5] || '')
      });
    }
    return { ok:true, origenes:out };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Alta/edición (upsert por id; gated por editar_ingresos) ───────────── */
function saveOrigen(b){
  try {
    if (!_tokenHasPermission((b && b.token) || '', 'editar_ingresos'))
      return { ok:false, error:'Sin autorizacion para editar origenes externos. Solicita el permiso editar_ingresos.' };

    var nombre = String((b && b.nombre) || '').trim();
    if (!nombre) return { ok:false, error:'El nombre del origen es requerido' };
    var tipo   = String((b && b.tipo) || '').trim() || 'Agencia';
    var alias  = String((b && b.alias) || '').trim();
    var notas  = String((b && b.notas) || '').trim();
    var activo = _origTruthy(b && b.activo);

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(ORIGENES_TAB);
    if (!sh){ setupOrigenesExternos(); sh = ss.getSheetByName(ORIGENES_TAB); }
    if (!sh) return { ok:false, error:'No se pudo crear Origenes_Externos' };

    var data = sh.getDataRange().getValues();
    var id = String((b && b.id) || '').trim();

    if (id){
      for (var i = 1; i < data.length; i++){
        if (String(data[i][0] || '').trim() === id){
          // Actualiza Nombre..Notas (cols 2-6) sin tocar CreadoEn (col 7).
          sh.getRange(i + 1, 2, 1, 5).setValues([[nombre, tipo, alias, activo, notas]]);
          return { ok:true, id:id, updated:true };
        }
      }
      // id no encontrado → cae a alta conservando el id recibido
    } else {
      id = 'ORIG-' + _origPad(_origMaxIdNum(data) + 1);
    }

    sh.appendRow([id, nombre, tipo, alias, activo, notas, new Date()]);
    return { ok:true, id:id, created:true };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Borrado (por rowNum; gated por editar_ingresos) ──────────────────── */
function deleteOrigen(b){
  try {
    if (!_tokenHasPermission((b && b.token) || '', 'editar_ingresos'))
      return { ok:false, error:'Sin autorizacion para borrar origenes externos. Solicita el permiso editar_ingresos.' };
    var rowNum = parseInt((b && b.rowNum), 10);
    if (!rowNum || rowNum < 2) return { ok:false, error:'rowNum invalido' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(ORIGENES_TAB);
    if (!sh) return { ok:false, error:'Origenes_Externos no encontrada' };
    if (rowNum > sh.getLastRow()) return { ok:false, error:'rowNum fuera de rango' };
    sh.deleteRow(rowNum);
    return { ok:true, deleted:rowNum };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Sugerencia histórica: propone un origen para cada ingreso externo del
   año que aún no tenga OrigenExterno asignado. Solo LECTURA. ──────────── */
function sugerirOrigenesHistorico(anio){
  try {
    anio = parseInt(anio, 10) || new Date().getFullYear();
    var reg = readOrigenes();
    var origenesArr = (reg && reg.origenes) ? reg.origenes : [];

    var ss = SpreadsheetApp.openById(_ingIdDeAnio(anio));
    var sh = null, all = ss.getSheets();
    for (var s = 0; s < all.length; s++){ if (all[s].getName() === BD_INGRESOS_TAB){ sh = all[s]; break; } }
    if (!sh) return { ok:false, error:'BD_Ingresos no encontrada para ' + anio };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok:true, anio:anio, total:0, sugerencias:[] };

    // Columna OrigenExterno por encabezado (puede no existir todavía).
    var hdr = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    var iOrig = -1;
    for (var c = 0; c < hdr.length; c++){ if (hdr[c].indexOf('origenexterno') > -1){ iOrig = c; break; } }

    function dt(v){ if (!v) return ''; if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth()+1).padStart(2,'0') + '-' + String(v.getDate()).padStart(2,'0'); return String(v); }
    function num(v){ if (typeof v === 'number') return v; var n = parseFloat(String(v||'').replace(/[$,\s]/g,'')); return isNaN(n)?0:n; }

    var sug = [];
    for (var i = 1; i < data.length; i++){
      var r = data[i];
      var op = String(r[0] || '').trim(); if (!op) continue;

      // Ya atribuido → se salta.
      var yaTiene = (iOrig > -1) ? String(r[iOrig] || '').trim() : '';
      if (yaTiene) continue;

      var categoria = String(r[4] || '');
      var producto  = String(r[5] || '');
      var pacReal   = String(r[3] || '');
      var ciclo     = String(r[20] || '');

      // Candidato externo: ciclo/categoría con 'extern' O es agencia.
      var esExterno = (_origNorm(ciclo).indexOf('extern') > -1) || (_origNorm(categoria).indexOf('extern') > -1);
      if (!esExterno && typeof _summaryEsAgencia === 'function'){
        try { esExterno = _summaryEsAgencia(categoria) || _summaryEsAgencia(producto) || _summaryEsAgencia(pacReal); } catch(eAg){}
      }
      if (!esExterno) continue;

      var matchText = pacReal + ' ' + categoria + ' ' + producto;
      var m = _origMatch(matchText, origenesArr);

      // Enmascara el paciente en la salida si el rol no ve datos sensibles
      // (el match usa el nombre real internamente).
      var pacShow = (typeof _privVer === 'function' && !_privVer() && typeof _privPaciente === 'function')
                    ? _privPaciente(op) : pacReal;

      sug.push({
        rowNum: i + 1,
        op: op,
        linea: num(r[1]),
        fecha: dt(r[2]),
        paciente: pacShow,
        categoria: categoria,
        producto: producto,
        total: num(r[9]),
        sugeridoNombre: m.nombre,
        sugeridoTipo: m.tipo,
        confianza: m.confianza
      });
    }
    return { ok:true, anio:anio, total:sug.length, sugerencias:sug };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Aplica la atribución histórica: escribe origenNombre en la columna
   OrigenExterno SOLO de los rowNums recibidos (gated por editar_ingresos). */
function aplicarOrigenesHistorico(b){
  try {
    if (!_tokenHasPermission((b && b.token) || '', 'editar_ingresos'))
      return { ok:false, error:'Sin autorizacion para atribuir origenes. Solicita el permiso editar_ingresos.' };

    var anio = parseInt((b && b.anio), 10) || new Date().getFullYear();
    var asigs = (b && b.asignaciones && b.asignaciones.length) ? b.asignaciones : [];
    if (!asigs.length) return { ok:true, aplicadas:0 };

    var ss = SpreadsheetApp.openById(_ingIdDeAnio(anio));
    var sh = null, all = ss.getSheets();
    for (var s = 0; s < all.length; s++){ if (all[s].getName() === BD_INGRESOS_TAB){ sh = all[s]; break; } }
    if (!sh) return { ok:false, error:'BD_Ingresos no encontrada para ' + anio };

    // Garantiza la columna OrigenExterno (append-safe, no desplaza nada).
    var oc = (typeof _ingColEnsure === 'function')
             ? _ingColEnsure(sh, 'origenexterno', 'OrigenExterno')
             : _origEnsureColLocal(sh);

    var lastRow = sh.getLastRow();
    var aplicadas = 0;
    for (var i = 0; i < asigs.length; i++){
      var rn = parseInt(asigs[i] && asigs[i].rowNum, 10);
      if (!rn || rn < 2 || rn > lastRow) continue;
      var val = String((asigs[i] && asigs[i].origenNombre) || '').trim();
      sh.getRange(rn, oc).setValue(val);
      aplicadas++;
    }

    try { CacheService.getScriptCache().remove('gas_ingresos_v1'); } catch(eC){}
    return { ok:true, aplicadas:aplicadas };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Fallback de ensure-column por si finance.gs no está en el despliegue. */
function _origEnsureColLocal(sh){
  var lastCol = sh.getLastColumn();
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); });
  for (var c = 0; c < hdrs.length; c++){ if (hdrs[c].indexOf('origenexterno') > -1) return c + 1; }
  sh.getRange(1, lastCol + 1).setValue('OrigenExterno');
  return lastCol + 1;
}
