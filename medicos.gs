/* ══════════════════════════════════════════════════════════════
   medicos.gs — Catálogo de MÉDICOS TRATANTES (por nombre + canal)
   --------------------------------------------------------------
   Padrón independiente de médicos tratantes: cada médico tiene un NOMBRE y
   pertenece a un CANAL (grupo: Agencia / Hestia / Externos). Alimenta el
   dropdown "Médico Tratante" de la ficha del paciente (por nombre individual,
   no por grupo). El CANAL y el ORIGEN son listas simples que se editan en
   Panel de Control → Formularios.

   Es un catálogo NUEVO y SEPARADO del de Orígenes Externos (ese se queda solo
   para la atribución de agencias y sus descuentos — no se toca). El vínculo con
   agencia/atribución ocurre por el Canal del médico (cuando es 'Agencia').

   Leer: cualquier sesión válida (llena el dropdown). Alta/edición/baja:
   permiso editar_pacientes (gate en finance.gs). Baja = soft-delete (Activo=NO),
   nunca borra la fila. Hoja "Medicos_Tratantes" en SHEET_ID.
   ══════════════════════════════════════════════════════════════ */

var MED_TAB = 'Medicos_Tratantes';
var MED_HEADERS = ['Id','Nombre','Canal','Activo','Notas','CreadoEn'];

function setupMedicos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(MED_TAB);
  if (!sh) sh = ss.insertSheet(MED_TAB);
  sh.clear();
  sh.getRange(1,1,1,MED_HEADERS.length).setValues([MED_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  sh.setFrozenRows(1);
  [90,260,140,70,300,160].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
  return {ok:true, tab:MED_TAB};
}
function _medSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(MED_TAB);
  if (!sh) { setupMedicos(); sh = ss.getSheetByName(MED_TAB); }
  return sh;
}
function _medNorm(s){ return String(s==null?'':s).trim(); }
function _medNormLow(s){ return _medNorm(s).toLowerCase(); }

// Lista de médicos. incluirInactivos=false → solo activos (para el dropdown).
function readMedicos(incluirInactivos) {
  try {
    var sh = _medSheet();
    var v = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < v.length; i++) {
      var r = v[i];
      var id = _medNorm(r[0]), nombre = _medNorm(r[1]);
      if (!id && !nombre) continue;
      var activo = (String(r[3]).toLowerCase() !== 'no' && String(r[3]).toLowerCase() !== 'false' && r[3] !== false);
      if (!incluirInactivos && !activo) continue;
      out.push({ rowNum:i+1, id:id, nombre:nombre, canal:_medNorm(r[2]), activo:activo, notas:_medNorm(r[4]) });
    }
    out.sort(function(a,b){ return a.nombre.localeCompare(b.nombre); });
    return { ok:true, medicos:out };
  } catch(ex){ return { ok:false, error:ex.message, medicos:[] }; }
}

// Alta o edición (upsert por Id; si no trae Id, alta con Id correlativo).
// Anti-duplicado por nombre (dentro de los activos). El dueño del gate valida
// el permiso editar_pacientes en el router.
function saveMedico(b) {
  try {
    b = b || {};
    var nombre = _medNorm(b.nombre);
    if (!nombre) return { ok:false, error:'El nombre del médico es obligatorio.' };
    var canal = _medNorm(b.canal);
    var notas = _medNorm(b.notas).substring(0, 500);
    var sh = _medSheet();
    var v = sh.getDataRange().getValues();
    var idIn = _medNorm(b.id);
    var rowNum = -1, max = 0;
    for (var i = 1; i < v.length; i++) {
      var rid = _medNorm(v[i][0]);
      var m = rid.match(/(\d+)/); if (m){ var n = parseInt(m[1],10); if (n>max) max=n; }
      if (idIn && rid === idIn) rowNum = i+1;
    }
    // Anti-duplicado por nombre (excluye la propia fila en edición).
    for (var j = 1; j < v.length; j++) {
      if ((j+1) === rowNum) continue;
      var act = (String(v[j][3]).toLowerCase() !== 'no' && String(v[j][3]).toLowerCase() !== 'false' && v[j][3] !== false);
      if (act && _medNormLow(v[j][1]) === nombre.toLowerCase())
        return { ok:false, error:'Ya existe un médico activo con el nombre "'+nombre+'".' };
    }
    var activo = (b.activo === undefined) ? true : (b.activo !== false && String(b.activo).toLowerCase() !== 'no');
    if (rowNum > 0) {
      // Update in-place SOLO de las columnas presentes (regla anti-blanqueo).
      sh.getRange(rowNum, 2).setValue(nombre);
      sh.getRange(rowNum, 3).setValue(canal);
      sh.getRange(rowNum, 4).setValue(activo);
      if (b.notas !== undefined) sh.getRange(rowNum, 5).setValue(notas);
      return { ok:true, id:idIn, actualizado:true };
    }
    var id = 'MED-' + String(max+1).padStart(4,'0');
    sh.appendRow([ id, nombre, canal, activo, notas, new Date() ]);
    return { ok:true, id:id, actualizado:false };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

// Baja = soft-delete (Activo=NO): deja registro, deja de aparecer en el dropdown.
function deleteMedico(b) {
  try {
    var id = _medNorm((b && b.id) || '');
    if (!id) return { ok:false, error:'Falta el id del médico.' };
    var sh = _medSheet();
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (_medNorm(v[i][0]) === id) { sh.getRange(i+1, 4).setValue(false); return { ok:true, id:id }; }
    }
    return { ok:false, error:'No se encontró el médico '+id+'.' };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
