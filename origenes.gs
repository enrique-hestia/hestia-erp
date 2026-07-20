/* ══════════════════════════════════════════════════════════════════════
   ORÍGENES EXTERNOS — atribución del "dueño" de un ingreso externo.

   Los ingresos del bucket "Externos" (CicloAltaBaja/Categoría con 'extern'
   o que matchea una AGENCIA) no tienen dueño estructurado: hoy se adivina
   por substrings. Este módulo agrega una dimensión de dueño ("Origen
   externo") con un CATÁLOGO editable (hoja Origenes_Externos) y un
   back-fill histórico asistido.

   JERARQUÍA (build .212): el catálogo dejó de ser plano. `Tipo` tiene 4
   valores (Grupo / Agencia / Médico externo / Coordinador) y la columna
   `Padre` amarra un origen a su Grupo — UN solo nivel (un médico pertenece a
   un grupo; un grupo no pertenece a nada). Con eso se puede contar los
   procedimientos de un GRUPO de médicos, no solo de uno.

   El mismo catálogo alimenta ahora la FICHA DEL PACIENTE (Canal / Médico
   Tratante, vía _origOpcionesFicha ← core.gs action=options): antes eran
   texto libre de la pestaña "Opciones", un tercer vocabulario que no conocía
   a los otros dos.

   Contrato con el frontend:
     GET  action=origenes                    → readOrigenes()  (trae padre+grupoId)
     GET  action=sugerirOrigenes&anio=YYYY    → sugerirOrigenesHistorico(anio)
     POST action=saveOrigen                   → saveOrigen(body)  (valida jerarquía)
     POST action=deleteOrigen                 → deleteOrigen(body)
     POST action=aplicarOrigenesHistorico     → aplicarOrigenesHistorico(body)

   Contrato con OTROS MÓDULOS (scope global):
     _origResolver(nombreOId) → {id, nombre, tipo, padreId, grupoId}
       grupoId = el padre si existe, o el propio id si no → "el grupo de un
       médico independiente es él mismo". Resuelve por ID, Nombre o Alias
       (exactos, normalizados); también los inactivos.
     _origResolverEn(catalogo, nombreOId) → igual, sin releer la hoja.

   Dependencias (mismo despliegue): SHEET_ID (config.gs), INGRESOS_SS_ID /
   _ingIdDeAnio / BD_INGRESOS_TAB / _tokenHasPermission / _ingColEnsure /
   PACIENTES_SS_ID (finance.gs), _summaryAgencias / _summaryEsAgencia
   (summary.gs), jsonResponse (core.gs).

   IMPORTANTE: correr setupOrigenesExternos() UNA vez desde el editor de
   Apps Script para crear/sembrar la hoja. La columna OrigenExterno en
   BD_Ingresos se crea sola en la primera escritura (ver _ingColEnsure).
   ══════════════════════════════════════════════════════════════════════ */

var ORIGENES_TAB     = 'Origenes_Externos';
/* `Padre` es APPEND-SAFE al final: las hojas viejas no la traen y se crea sola
   en la primera escritura (_origColEnsure). NUNCA se lee por posición. */
var ORIGENES_HEADERS = ['ID','Nombre','Tipo','Alias','Activo','Notas','CreadoEn','Padre'];

/* Tipos canónicos del catálogo. Antes solo existían 'Agencia' y 'Médico externo'
   y la vista los separaba por substring ('agenc' → agencia, TODO lo demás →
   médico). Con 4 tipos ese criterio ya no alcanza: 'Coordinador' y 'Grupo'
   caerían en "médico externo" y la jerarquía se vería al revés. */
var ORIG_TIPOS = ['Grupo','Agencia','Médico externo','Coordinador'];

/* Tipo canónico a partir del texto de la celda. Tolerante (los datos ya
   escritos dicen 'Agencia' / 'Médico externo' y DEBEN seguir clasificando
   igual). El orden importa: 'agenc' antes que nada, y 'medic' al final para
   que "Grupo Médico" escrito como tipo no se lea como médico.
   Celda vacía → 'Médico externo', que es como los clasificaba la vista vieja
   (todo lo que no era agencia). */
function _origTipoCanon(tipo){
  var n = _origNorm(tipo);
  if (!n) return 'Médico externo';
  if (n.indexOf('agenc')   > -1) return 'Agencia';
  if (n.indexOf('coordin') > -1) return 'Coordinador';
  if (n.indexOf('grupo')   > -1) return 'Grupo';
  if (n.indexOf('medic')   > -1 || n.indexOf('doctor') > -1) return 'Médico externo';
  return 'Médico externo';
}
/* ¿Este tipo puede ser PADRE de otro? Solo los Grupos. Un solo nivel. */
function _origTipoEsGrupo(tipo){ return _origTipoCanon(tipo) === 'Grupo'; }

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

/* ── Columnas POR ENCABEZADO (nunca por posición) ──────────────────────
   Mismo patrón que _ingColEnsure/_egColEnsure: busca el header normalizado;
   si no está, la crea al FINAL y devuelve su índice. Así una hoja vieja
   (7 columnas, sin Padre) se migra sola sin desplazar nada. */
function _origColFind(hdrRow, key){
  if (!hdrRow) return -1;
  for (var c = 0; c < hdrRow.length; c++){ if (_origNorm(hdrRow[c]) === key) return c; }
  return -1;
}
function _origColEnsure(sh, key, headerName){
  var lastCol = sh.getLastColumn();
  var hdrs = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var i = _origColFind(hdrs, key);
  if (i > -1) return i + 1;
  sh.getRange(1, lastCol + 1).setValue(headerName);
  return lastCol + 1;
}

/* Índices de TODAS las columnas del catálogo, por encabezado, con fallback a
   las posiciones históricas (0..6) si la hoja no trae encabezados legibles.
   Padre → -1 si la hoja todavía no la tiene (se trata como vacía). */
function _origIdx(hdrRow){
  function f(key, fb){ var i = _origColFind(hdrRow, key); return i > -1 ? i : fb; }
  return { id:f('id',0), nombre:f('nombre',1), tipo:f('tipo',2), alias:f('alias',3),
           activo:f('activo',4), notas:f('notas',5), creado:f('creadoen',6),
           padre:_origColFind(hdrRow, 'padre') };
}

/* ── Construye los registros del catálogo desde la matriz cruda, resolviendo
   la jerarquía. PURA (sin SpreadsheetApp) → es la que se prueba. ────────── */
function _origBuild(data){
  var out = [];
  if (!data || data.length < 2) return out;
  var ix = _origIdx(data[0]);
  for (var i = 1; i < data.length; i++){
    var r = data[i] || [];
    if (!String(r[ix.id] || '').trim() && !String(r[ix.nombre] || '').trim()) continue;
    out.push({
      rowNum: i + 1,
      id:     String(r[ix.id] || ''),
      nombre: String(r[ix.nombre] || ''),
      tipo:   _origTipoCanon(r[ix.tipo]),
      alias:  String(r[ix.alias] || ''),
      activo: _origTruthy(r[ix.activo]),
      notas:  String(r[ix.notas] || ''),
      padre:  (ix.padre > -1) ? String(r[ix.padre] || '').trim() : ''
    });
  }
  return _origResolveGrupos(out);
}

/* Rellena `padre` (limpiando referencias colgantes) y `grupoId` en cada
   registro. grupoId = el padre si existe, o el propio id si no → "el grupo de
   un médico independiente es él mismo". Defensivo contra datos sucios: un
   padre que no existe, que se apunta a sí mismo, o que NO es Grupo (jerarquía
   de 2 niveles metida a mano en la hoja) se ignora → el registro queda suelto
   en vez de mentir sobre su grupo. */
function _origResolveGrupos(arr){
  var byId = {};
  for (var i = 0; i < arr.length; i++){
    var k = String(arr[i].id || '').trim();
    if (k) byId[k] = arr[i];
  }
  for (var j = 0; j < arr.length; j++){
    var o  = arr[j];
    var id = String(o.id || '').trim();
    var p  = String(o.padre || '').trim();
    if (p === id) p = '';                                  // ciclo directo
    var pr = p ? byId[p] : null;
    if (!pr) p = '';                                       // padre colgante
    else if (!_origTipoEsGrupo(pr.tipo)) p = '';           // el padre no es Grupo
    else if (String(pr.padre || '').trim()) p = '';        // el padre tiene padre → 2 niveles
    if (_origTipoEsGrupo(o.tipo)) p = '';                  // un Grupo no tiene padre
    o.padre   = p;
    o.grupoId = p || id;
  }
  return arr;
}

/* ── RESOLVER PÚBLICO (contrato compartido) ───────────────────────────────
   _origResolver(nombreOId) → {id, nombre, tipo, padreId, grupoId}
   Busca por ID exacto, luego por Nombre normalizado exacto, luego por Alias
   exacto. NO usa substring (para eso está _origMatch, que es difuso y sirve
   para adivinar en texto libre). Devuelve el objeto vacío si no existe.
   Resuelve también los INACTIVOS: un nombre histórico debe seguir resolviendo
   aunque el origen ya se haya dado de baja. */
function _origResolver(nombreOId){
  var reg;
  try { reg = readOrigenes(); } catch(ex){ reg = null; }
  return _origResolverEn((reg && reg.origenes) ? reg.origenes : [], nombreOId);
}
/* Versión pura (recibe el catálogo ya leído) — la que se prueba y la que usa
   cualquier bucle que no quiera releer la hoja por fila. */
function _origResolverEn(arr, nombreOId){
  var vacio = { id:'', nombre:'', tipo:'', padreId:'', grupoId:'' };
  var q = String(nombreOId == null ? '' : nombreOId).trim();
  if (!q || !arr || !arr.length) return vacio;
  var qn = _origNorm(q);
  var hit = null;
  for (var i = 0; i < arr.length && !hit; i++){
    if (String(arr[i].id || '').trim() === q) hit = arr[i];
  }
  for (var j = 0; j < arr.length && !hit; j++){
    if (qn && _origNorm(arr[j].nombre) === qn) hit = arr[j];
  }
  for (var k = 0; k < arr.length && !hit; k++){
    var als = String(arr[k].alias || '').split(',');
    for (var a = 0; a < als.length; a++){
      var an = _origNorm(als[a]);
      if (an && an === qn){ hit = arr[k]; break; }
    }
  }
  if (!hit) return vacio;
  // grupoId puede no venir pre-resuelto si el arreglo no pasó por _origBuild.
  var id = String(hit.id || '').trim();
  var padreId = String(hit.padre || '').trim();
  var gid = String(hit.grupoId || '').trim() || padreId || id;
  return { id:id, nombre:String(hit.nombre || ''), tipo:_origTipoCanon(hit.tipo),
           padreId:padreId, grupoId:gid };
}

/* ── Opciones para la FICHA DEL PACIENTE (las consume core.gs action=options).
   Canal  → de dónde viene: Grupos y Agencias.
   Médico → Médicos externos (el frontend los filtra por grupo del Canal).
   ok:false si el catálogo está vacío → core.gs conserva la hoja Opciones y no
   deja la ficha sin opciones por un catálogo a medio sembrar. */
function _origOpcionesFicha(){
  var reg;
  try { reg = readOrigenes(); } catch(ex){ return { ok:false }; }
  if (!reg || !reg.ok) return { ok:false };
  var act = [], all = reg.origenes || [];
  for (var i = 0; i < all.length; i++){ if (all[i].activo !== false) act.push(all[i]); }
  if (!act.length) return { ok:false };
  var canales = [], medicos = [], registros = [];
  for (var j = 0; j < act.length; j++){
    var o = act[j], t = o.tipo;
    if (t === 'Grupo' || t === 'Agencia') canales.push(o.nombre);
    if (t === 'Médico externo') medicos.push(o.nombre);
    registros.push({ id:o.id, nombre:o.nombre, tipo:t, padre:o.padre || '', grupoId:o.grupoId || o.id });
  }
  function az(a,b){ return _origNorm(a) < _origNorm(b) ? -1 : (_origNorm(a) > _origNorm(b) ? 1 : 0); }
  canales.sort(az); medicos.sort(az);
  return { ok:true, canales:canales, medicos:medicos, registros:registros };
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
      // El último '' es Padre: una agencia sembrada no pertenece a ningún Grupo.
      toAdd.push(['ORIG-' + _origPad(counter), nom, 'Agencia', '', true, 'Sembrado desde AGENCIAS', new Date(), '']);
      existing[_origNorm(nom)] = true;
    }
    if (toAdd.length) sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, ORIGENES_HEADERS.length).setValues(toAdd);

    return { ok:true, created:created, seeded:toAdd.length, tab:ORIGENES_TAB };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Lectura del catálogo (devuelve `padre` y `grupoId` — contrato) ────── */
function readOrigenes(){
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(ORIGENES_TAB);
    if (!sh) return { ok:true, origenes:[] };
    return { ok:true, origenes:_origBuild(sh.getDataRange().getValues()) };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Validación de la JERARQUÍA (pura → se prueba). Recibe el catálogo ya
   construido, el id que se está guardando ('' si es alta), el tipo canónico
   y el padre propuesto. Devuelve '' si es válido, o el mensaje de error.
   Reglas:
     · Un Grupo no puede tener padre (la jerarquía es de UN nivel).
     · El padre debe existir.
     · El padre no puede ser uno mismo (ciclo directo).
     · El padre debe ser de tipo Grupo → esto por sí solo impide 2 niveles,
       porque un Grupo nunca tiene padre. Se valida igual de forma explícita
       por si la hoja trae datos sucios metidos a mano.
     · Dejar de ser Grupo teniendo miembros colgando huerfanaría a los hijos. */
function _origValidaJerarquia(arr, id, tipoCanon, padreId){
  arr = arr || [];
  id = String(id || '').trim();
  padreId = String(padreId || '').trim();

  if (tipoCanon === 'Grupo' && padreId)
    return 'Un Grupo no puede pertenecer a otro origen: la jerarquia es de un solo nivel.';

  if (id && tipoCanon !== 'Grupo'){
    var hijos = 0;
    for (var h = 0; h < arr.length; h++){
      if (String(arr[h].padre || '').trim() === id) hijos++;
    }
    if (hijos)
      return 'No puedes cambiar el tipo: ' + hijos + ' origen(es) tienen a este como Grupo. Muevelos primero.';
  }

  if (!padreId) return '';
  if (padreId === id) return 'Un origen no puede ser su propio Grupo.';

  var pr = null;
  for (var i = 0; i < arr.length; i++){
    if (String(arr[i].id || '').trim() === padreId){ pr = arr[i]; break; }
  }
  if (!pr) return 'El Grupo padre indicado (' + padreId + ') no existe en el catalogo.';
  if (!_origTipoEsGrupo(pr.tipo))
    return 'El padre debe ser de tipo Grupo. "' + String(pr.nombre || '') + '" es ' + _origTipoCanon(pr.tipo) + '.';
  if (String(pr.padre || '').trim())
    return 'El Grupo "' + String(pr.nombre || '') + '" ya pertenece a otro: no se permiten jerarquias de mas de un nivel.';
  return '';
}

/* ── Alta/edición (upsert por id; gated por editar_ingresos) ───────────── */
function saveOrigen(b){
  try {
    if (!_tokenHasPermission((b && b.token) || '', 'editar_ingresos'))
      return { ok:false, error:'Sin autorizacion para editar origenes externos. Solicita el permiso editar_ingresos.' };

    var nombre = String((b && b.nombre) || '').trim();
    if (!nombre) return { ok:false, error:'El nombre del origen es requerido' };
    var tipo   = _origTipoCanon(String((b && b.tipo) || '').trim() || 'Agencia');
    var alias  = String((b && b.alias) || '').trim();
    var notas  = String((b && b.notas) || '').trim();
    var activo = _origTruthy(b && b.activo);
    var padre  = String((b && b.padre) || '').trim();

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(ORIGENES_TAB);
    if (!sh){ setupOrigenesExternos(); sh = ss.getSheetByName(ORIGENES_TAB); }
    if (!sh) return { ok:false, error:'No se pudo crear Origenes_Externos' };

    // La columna Padre se garantiza ANTES de leer, para que _origBuild la vea.
    var cPadre = _origColEnsure(sh, 'padre', 'Padre');

    var data = sh.getDataRange().getValues();
    var id   = String((b && b.id) || '').trim();

    // Validación contra el catálogo REAL (con la jerarquía ya resuelta).
    var err = _origValidaJerarquia(_origBuild(data), id, tipo, padre);
    if (err) return { ok:false, error:err };
    if (tipo === 'Grupo') padre = '';

    var ix = _origIdx(data[0] || []);
    if (id){
      for (var i = 1; i < data.length; i++){
        if (String(data[i][ix.id] || '').trim() === id){
          // Escritura POR COLUMNA (no por bloque fijo): la hoja puede traer
          // columnas nuevas de por medio y CreadoEn no se toca.
          sh.getRange(i + 1, ix.nombre + 1).setValue(nombre);
          sh.getRange(i + 1, ix.tipo   + 1).setValue(tipo);
          sh.getRange(i + 1, ix.alias  + 1).setValue(alias);
          sh.getRange(i + 1, ix.activo + 1).setValue(activo);
          sh.getRange(i + 1, ix.notas  + 1).setValue(notas);
          sh.getRange(i + 1, cPadre).setValue(padre);
          return { ok:true, id:id, updated:true };
        }
      }
      // id no encontrado → cae a alta conservando el id recibido
    } else {
      id = 'ORIG-' + _origPad(_origMaxIdNum(data) + 1);
    }

    var fila = [];
    fila[ix.id] = id; fila[ix.nombre] = nombre; fila[ix.tipo] = tipo;
    fila[ix.alias] = alias; fila[ix.activo] = activo; fila[ix.notas] = notas;
    fila[ix.creado] = new Date(); fila[cPadre - 1] = padre;
    var ancho = Math.max(sh.getLastColumn(), cPadre);
    for (var f = 0; f < ancho; f++){ if (fila[f] === undefined) fila[f] = ''; }
    sh.getRange(sh.getLastRow() + 1, 1, 1, ancho).setValues([fila.slice(0, ancho)]);
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

    // Borrar un Grupo con miembros los dejaría apuntando a un padre inexistente
    // (_origResolveGrupos los soltaría en silencio y el reporte por grupo
    // cambiaría sin que nadie se entere). Se bloquea con un mensaje accionable.
    var arr = _origBuild(sh.getDataRange().getValues());
    var yo = null;
    for (var i = 0; i < arr.length; i++){ if (arr[i].rowNum === rowNum){ yo = arr[i]; break; } }
    if (yo && _origTipoEsGrupo(yo.tipo)){
      var hijos = [];
      for (var j = 0; j < arr.length; j++){
        if (String(arr[j].padre || '').trim() === String(yo.id || '').trim()) hijos.push(arr[j].nombre);
      }
      if (hijos.length)
        return { ok:false, error:'No se puede borrar el Grupo "' + yo.nombre + '": lo integran ' + hijos.length
                 + ' origen(es) (' + hijos.slice(0,4).join(', ') + (hijos.length>4?'…':'')
                 + '). Sacalos del grupo primero.' };
    }

    sh.deleteRow(rowNum);
    return { ok:true, deleted:rowNum };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ══════════════════════════════════════════════════════════════════════
   MIGRACIÓN: Canal / Médico Tratante de la ficha → catálogo de orígenes.
   ══════════════════════════════════════════════════════════════════════
   La ficha del paciente guarda TEXTO LIBRE que salía de la pestaña
   "Opciones". Ese texto y el catálogo Origenes_Externos nunca se conocieron:
   por eso no se puede reportar por médico aunque el dato "esté".

   SOLO LECTURA por default. Devuelve el mapeo completo para que el usuario lo
   vea ANTES de que se escriba nada. Con {aplicar:true} normaliza en la ficha
   los valores que resuelven de forma EXACTA (nombre o alias) al nombre
   canónico del catálogo — nada más. Es idempotente: si ya está canónico, no
   escribe. Los valores que NO empatan NO se borran ni se tocan: se reportan.

   Uso desde el editor de Apps Script:
     migrarCanalesAOrigenes()                        → reporte
     migrarCanalesAOrigenes({aplicar:true, token:…}) → normaliza
   ══════════════════════════════════════════════════════════════════════ */
function migrarCanalesAOrigenes(opts){
  try {
    opts = opts || {};
    var aplicar = (opts.aplicar === true);

    // La escritura pide permiso; la lectura del reporte no (no revela dinero).
    if (aplicar && !_tokenHasPermission(String(opts.token || ''), 'editar_ingresos'))
      return { ok:false, error:'Sin autorizacion para escribir en las fichas. Solicita el permiso editar_ingresos.' };

    var reg = readOrigenes();
    if (!reg || !reg.ok) return { ok:false, error:'No se pudo leer el catalogo: ' + ((reg && reg.error) || '') };
    var cat = reg.origenes || [];
    if (!cat.length) return { ok:false, error:'El catalogo Origenes_Externos esta vacio. Corre setupOrigenesExternos() y da de alta los grupos/medicos primero.' };

    var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
    var sh = ss.getSheets()[0];
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok:true, pacientes:0, campos:[] };

    var hdr = data[0];
    function colDe(nombreCol){
      var target = _origNorm(nombreCol);
      for (var c = 0; c < hdr.length; c++){ if (_origNorm(hdr[c]) === target) return c; }
      return -1;
    }
    var campos = [
      { campo:'Canal',           col:colDe('Canal'),           tiposOk:['Grupo','Agencia'] },
      { campo:'Médico Tratante', col:colDe('Medico Tratante'), tiposOk:['Médico externo'] }
    ];

    var out = { ok:true, aplicar:aplicar, pacientes:data.length - 1, hoja:sh.getName(), campos:[] };
    var escrituras = [];

    for (var f = 0; f < campos.length; f++){
      var cf = campos[f];
      var res = { campo:cf.campo, columna:cf.col + 1, encontrada:(cf.col > -1),
                  valores:[], empatan:0, noEmpatan:0, filasConValor:0, normalizables:0 };
      if (cf.col < 0){ res.error = 'La columna "' + cf.campo + '" no existe en la hoja de Pacientes.'; out.campos.push(res); continue; }

      // Agrupa los valores DISTINTOS (normalizados) con su conteo y sus filas.
      var mapa = {};
      for (var i = 1; i < data.length; i++){
        var raw = String(data[i][cf.col] == null ? '' : data[i][cf.col]).trim();
        if (!raw) continue;
        res.filasConValor++;
        var k = _origNorm(raw);
        if (!mapa[k]) mapa[k] = { valor:raw, veces:0, filas:[] };
        mapa[k].veces++;
        mapa[k].filas.push(i + 1);
      }

      var claves = Object.keys(mapa);
      for (var v = 0; v < claves.length; v++){
        var e = mapa[claves[v]];
        var r = _origResolverEn(cat, e.valor);          // exacto: id / nombre / alias
        var item = { valor:e.valor, veces:e.veces, filas:e.filas.slice(0, 12), totalFilas:e.filas.length };

        if (r.id){
          item.estado    = 'empata';
          item.origenId  = r.id;
          item.canonico  = r.nombre;
          item.tipo      = r.tipo;
          item.grupoId   = r.grupoId;
          item.tipoEsperadoOk = (cf.tiposOk.indexOf(r.tipo) > -1);
          if (!item.tipoEsperadoOk)
            item.aviso = 'Resuelve a un origen de tipo "' + r.tipo + '", que no es de los que van en ' + cf.campo + '.';
          // ¿Hay que normalizar la escritura? Solo si el texto difiere del canónico.
          item.normalizar = (e.valor !== r.nombre);
          if (item.normalizar){
            res.normalizables++;
            if (aplicar){
              for (var ff = 0; ff < e.filas.length; ff++)
                escrituras.push({ fila:e.filas[ff], col:cf.col + 1, valor:r.nombre });
            }
          }
          res.empatan++;
        } else {
          // Sin match exacto → se intenta el difuso SOLO como pista para el
          // usuario. NUNCA se escribe: adivinar aquí es corromper 362 fichas.
          var m = _origMatch(e.valor, cat);
          item.estado    = 'sin_match';
          item.sugerencia = m.nombre || '';
          item.sugerenciaTipo = m.tipo || '';
          item.confianza = m.confianza || '';
          item.accion = m.nombre
            ? 'Se parece a "' + m.nombre + '". Confirmalo a mano (o agregale ese alias al origen).'
            : 'Da de alta este origen en el catalogo, o corrige la ficha.';
          res.noEmpatan++;
        }
        res.valores.push(item);
      }
      res.valores.sort(function(a,b){ return b.veces - a.veces; });
      out.campos.push(res);
    }

    out.escriturasPendientes = escrituras.length;
    if (aplicar && escrituras.length){
      for (var w = 0; w < escrituras.length; w++)
        sh.getRange(escrituras[w].fila, escrituras[w].col).setValue(escrituras[w].valor);
      out.escritas = escrituras.length;
    } else {
      out.escritas = 0;
      if (!aplicar) out.nota = 'REPORTE (no se escribio nada). Para normalizar los valores que empatan: migrarCanalesAOrigenes({aplicar:true, token:"<tu token>"}).';
    }
    return out;
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

      // Filtro por AÑO de la fila. BD_Ingresos suele estar CONSOLIDADO (todos los
      // años en un solo libro) y _ingIdDeAnio devuelve el mismo libro para
      // cualquier año → sin esto se colaban filas de otros años (ej. 2024
      // apareciendo con el filtro en 2026). Solo se descarta cuando el año es
      // CONOCIDO y distinto; fechas ilegibles no se ocultan.
      var fRow = r[2];
      var yRow = (fRow instanceof Date) ? fRow.getFullYear() : parseInt(String(dt(fRow)).substring(0,4), 10);
      if (yRow && yRow !== anio) continue;

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
