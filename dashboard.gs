/* ══════════════════════════════════════════════════════════════
   dashboard.gs — Tableros configurables ("Mi tablero") por usuario
   --------------------------------------------------------------
   Persiste el layout del Resumen configurable de cada usuario (qué widgets,
   tipo, tamaño, orden) para que sea igual en todos sus dispositivos. El
   frontend usa localStorage como caché instantáneo y sincroniza aquí.

   1 fila = 1 layout completo; los widgets van como blob JSON en ConfigJSON.
   El DUEÑO se deriva del TOKEN (verifyToken en el router), NUNCA del cliente,
   para que nadie pueda leer ni sobrescribir el tablero de otro.
   Upsert IN-PLACE por (Usuario, LayoutId) — NUNCA sh.clear() (hoja multiusuario;
   borrar-y-reescribir perdería los tableros de todos los demás).
   Hoja "Dashboard_Layouts" en SHEET_ID.
   ══════════════════════════════════════════════════════════════ */

var DASH_TAB = 'Dashboard_Layouts';
var DASH_HEADERS = ['Usuario','LayoutId','Nombre','EsPlantilla','Orden','Activo','ConfigJSON','ActualizadoEn'];

function setupDashboard() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(DASH_TAB);
  if (!sh) sh = ss.insertSheet(DASH_TAB);
  sh.clear();
  sh.getRange(1,1,1,DASH_HEADERS.length).setValues([DASH_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#c46a7a');
  sh.setFrozenRows(1);
  [230,140,180,90,70,70,520,160].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
  return {ok:true, tab:DASH_TAB};
}

function _dashSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(DASH_TAB);
  if (!sh) { setupDashboard(); sh = ss.getSheetByName(DASH_TAB); }
  return sh;
}
function _dashNorm(s){ return String(s==null?'':s).toLowerCase().trim(); }

// Lee los layouts del usuario (dueño = email del token) + las plantillas globales.
function readDashboardLayouts(usuario) {
  try {
    usuario = _dashNorm(usuario);
    if (!usuario) return { ok:false, error:'Sesión inválida.', layouts:[], plantillas:[] };
    var sh = _dashSheet();
    var v = sh.getDataRange().getValues();
    var mine=[], globales=[];
    for (var i=1;i<v.length;i++){
      var r=v[i]; var u=_dashNorm(r[0]); var lid=String(r[1]||'').trim();
      if (!lid) continue;
      var esPl = (r[3]===true || String(r[3]).toLowerCase()==='true');
      var cfg = {};
      try { cfg = r[6] ? JSON.parse(r[6]) : {}; } catch(e){ cfg = {}; }
      var item = { layoutId:lid, nombre:String(r[2]||''), esPlantilla:esPl,
                   orden:Number(r[4]||0), activo:(r[5]===true||String(r[5]).toLowerCase()==='true'),
                   config:cfg };
      if (u === usuario) mine.push(item);
      else if (u === '__global__' || esPl) globales.push(item);
    }
    return { ok:true, layouts:mine, plantillas:globales };
  } catch(ex){ return { ok:false, error:ex.message, layouts:[], plantillas:[] }; }
}

// Guarda (upsert) un layout PERSONAL. El dueño SIEMPRE es el email del token:
// un usuario no puede escribir el tablero de otro ni marcar plantillas globales
// (esas se siembran con seed/setup, no por esta vía).
function saveDashboardLayout(usuario, layout) {
  try {
    usuario = _dashNorm(usuario);
    if (!usuario) return { ok:false, error:'Sesión inválida.' };
    layout = layout || {};
    var lid = String(layout.layoutId||'default').trim() || 'default';
    var cfgJson = JSON.stringify(layout.config || {widgets:[]});
    var sh = _dashSheet();
    var v = sh.getDataRange().getValues();
    var rowNum = -1;
    for (var i=1;i<v.length;i++){
      if (_dashNorm(v[i][0])===usuario && String(v[i][1]||'').trim()===lid){ rowNum=i+1; break; }
    }
    var rowVals = [ usuario, lid, String(layout.nombre||'Mi tablero'), false,
                    Number(layout.orden||0), (layout.activo!==false), cfgJson, new Date() ];
    if (rowNum>0) sh.getRange(rowNum,1,1,DASH_HEADERS.length).setValues([rowVals]);
    else sh.appendRow(rowVals);
    return { ok:true, layoutId:lid };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

// Borra un layout del usuario (por si maneja varios). Solo puede borrar el suyo.
function deleteDashboardLayout(usuario, layoutId) {
  try {
    usuario=_dashNorm(usuario); var lid=String(layoutId||'').trim();
    if(!usuario||!lid) return {ok:false,error:'Faltan datos.'};
    var sh=_dashSheet(); var v=sh.getDataRange().getValues();
    for (var i=v.length-1;i>=1;i--){
      if(_dashNorm(v[i][0])===usuario && String(v[i][1]||'').trim()===lid){ sh.deleteRow(i+1); return {ok:true}; }
    }
    return {ok:false, error:'No se encontró el tablero.'};
  } catch(ex){ return {ok:false, error:ex.message}; }
}
