/* ═══════════════════════════════════════════════════════════════════════════
   EXPEDIENTE DE MÉDICOS TRATANTES
   Documentación que la clínica exige a cada médico externo para poder operar
   (requisito de salud). Checklist configurable, archivos en Drive, vigencias y
   avisos/notificaciones (manual o nocturna) al médico para que renueve.

   Los médicos VIVEN en Origenes_Externos (tipo 'Médico externo'); aquí solo se
   guarda su EXPEDIENTE. NO hay datos de pacientes en este módulo.

   Hojas (en SHEET_ID):
     · Config_Expediente : DocId | Nombre | RequiereVigencia | DiasAviso | Orden | Activo
     · Expedientes_Medicos : Medico | DocId | DriveUrl | FileId | FechaSubida | FechaVigencia | Usuario | ActualizadoEn
   Drive: carpeta raíz "Expedientes Médicos" (Script Property EXPEDIENTE_DRIVE_ID),
     una subcarpeta por médico.
   Requiere: finance.gs (_tokenHasPermission, logAudit), recordatorios.gs (saveRecordatorio),
     comprobantes/GmailApp para el correo. Permiso de escritura: 'editar_ingresos'.
   ═══════════════════════════════════════════════════════════════════════════ */

var EXPEDIENTE_VER      = 'expediente-2026.07.21a';
var EXPEDIENTE_CFG_TAB  = 'Config_Expediente';
var EXPEDIENTE_TAB      = 'Expedientes_Medicos';
var EXPEDIENTE_PERM     = 'editar_ingresos';           // alta/edición del expediente
var EXPEDIENTE_DRIVE_KEY= 'EXPEDIENTE_DRIVE_ID';

// Documentos que se piden por defecto (se siembran en Config_Expediente; luego el
// admin los edita en Panel de Control). requiereVigencia = documento que caduca.
var EXPEDIENTE_DEFAULTS = [
  { docId:'cv',            nombre:'Curriculum Vitae resumido',                      vig:false, dias:0  },
  { docId:'titulo_med',    nombre:'Título de Medicina',                             vig:false, dias:0  },
  { docId:'titulo_esp',    nombre:'Título de Especialidad en Ginecología',          vig:false, dias:0  },
  { docId:'cedula_prof',   nombre:'Cédula Profesional',                             vig:false, dias:0  },
  { docId:'cedula_esp',    nombre:'Cédula de Especialidad',                          vig:false, dias:0  },
  { docId:'cert_consejo',  nombre:'Certificado del Consejo Mexicano de Ginecología', vig:true,  dias:60 },
  { docId:'ine',           nombre:'Identificación oficial',                         vig:false, dias:0  },
  { docId:'poliza_rc',     nombre:'Póliza de Responsabilidad Civil',                vig:true,  dias:30 },
  { docId:'csf',           nombre:'Constancia de Situación Fiscal',                 vig:false, dias:0  }
];

function _expSS() { return SpreadsheetApp.openById(SHEET_ID); }
function _expStr(v){ if(v instanceof Date && !isNaN(v)) return v.getFullYear()+'-'+('0'+(v.getMonth()+1)).slice(-2)+'-'+('0'+v.getDate()).slice(-2); return String(v==null?'':v).trim(); }
function _expKeyMed(n){ return String(n||'').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,''); }
// Anti-inyección de fórmulas al escribir texto de usuario en una hoja.
function _expCell(v){ v=String(v==null?'':v); return /^[=+\-@\t\r]/.test(v) ? "'"+v : v; }

/* ── Config: qué documentos se piden ─────────────────────────────────────── */
function _expEnsureCfg() {
  var ss = _expSS(); var sh = ss.getSheetByName(EXPEDIENTE_CFG_TAB);
  if (!sh) {
    sh = ss.insertSheet(EXPEDIENTE_CFG_TAB);
    sh.appendRow(['DocId','Nombre','RequiereVigencia','DiasAviso','Orden','Activo']);
    sh.setFrozenRows(1); sh.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f4f6');
    var rows = EXPEDIENTE_DEFAULTS.map(function(d,i){ return [d.docId, d.nombre, d.vig, d.dias, (i+1)*10, true]; });
    sh.getRange(2,1,rows.length,6).setValues(rows);
  }
  return sh;
}
function readExpedienteConfig() {
  try {
    var sh = _expEnsureCfg(); var raw = sh.getDataRange().getValues(); var out = [];
    for (var i=1;i<raw.length;i++){
      if (!String(raw[i][0]||'').trim()) continue;
      out.push({ docId:String(raw[i][0]).trim(), nombre:String(raw[i][1]||'').trim(),
        requiereVigencia: raw[i][2]===true || String(raw[i][2]).toLowerCase()==='true' || String(raw[i][2]).toLowerCase()==='si' || String(raw[i][2]).toLowerCase()==='sí',
        diasAviso: parseInt(raw[i][3],10)||0, orden: parseInt(raw[i][4],10)||999,
        activo: !(raw[i][5]===false || String(raw[i][5]).toLowerCase()==='false' || String(raw[i][5]).toLowerCase()==='no') });
    }
    out.sort(function(a,b){ return a.orden-b.orden; });
    return { ok:true, version:EXPEDIENTE_VER, docs:out };
  } catch(ex){ return { ok:false, error:ex.message, docs:[] }; }
}
function saveExpedienteConfig(body) {
  try {
    body = body||{};
    if (typeof _tokenHasPermission!=='function' || !_tokenHasPermission(body.token, EXPEDIENTE_PERM))
      return { ok:false, error:'No tienes permiso para configurar el expediente ('+EXPEDIENTE_PERM+').' };
    var docs = body.docs||[]; var sh = _expEnsureCfg();
    var last = sh.getLastRow(); if (last>1) sh.getRange(2,1,last-1,6).clearContent();
    if (docs.length) {
      var rows = docs.map(function(d,i){ return [_expCell(d.docId||('doc'+i)), _expCell(d.nombre||''), !!d.requiereVigencia, parseInt(d.diasAviso,10)||0, parseInt(d.orden,10)||((i+1)*10), d.activo!==false]; });
      sh.getRange(2,1,rows.length,6).setValues(rows);
    }
    try { logAudit(body.usuario||'sistema','Expediente','GuardarConfig','','','',docs.length+' docs'); } catch(e){}
    return { ok:true, version:EXPEDIENTE_VER, guardados:docs.length };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Registros del expediente por médico ─────────────────────────────────── */
function _expEnsure() {
  var ss = _expSS(); var sh = ss.getSheetByName(EXPEDIENTE_TAB);
  if (!sh) {
    sh = ss.insertSheet(EXPEDIENTE_TAB);
    sh.appendRow(['Medico','DocId','DriveUrl','FileId','FechaSubida','FechaVigencia','Usuario','ActualizadoEn']);
    sh.setFrozenRows(1); sh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#f3f4f6');
  }
  return sh;
}
function _expLeerTodos() {
  var sh = _expEnsure(); var raw = sh.getDataRange().getValues(); var out = [];
  for (var i=1;i<raw.length;i++){
    if (!String(raw[i][0]||'').trim()) continue;
    out.push({ rowNum:i+1, medico:String(raw[i][0]).trim(), medicoKey:_expKeyMed(raw[i][0]), docId:String(raw[i][1]||'').trim(),
      driveUrl:String(raw[i][2]||'').trim(), fileId:String(raw[i][3]||'').trim(), fechaSubida:_expStr(raw[i][4]),
      fechaVigencia:_expStr(raw[i][5]), usuario:String(raw[i][6]||'').trim() });
  }
  return out;
}

/* Estado de vigencia de un documento (hoy vs FechaVigencia + diasAviso). */
function _expEstadoVig(fechaVig, diasAviso) {
  if (!fechaVig) return { estado:'sin_vigencia' };
  var hoy = new Date(); hoy.setHours(0,0,0,0);
  var fv = new Date(fechaVig+'T00:00:00'); if (isNaN(fv)) return { estado:'sin_vigencia' };
  var dias = Math.round((fv.getTime()-hoy.getTime())/86400000);
  if (dias < 0)  return { estado:'vencido',  dias:dias };
  if (dias <= (diasAviso||30)) return { estado:'por_vencer', dias:dias };
  return { estado:'vigente', dias:dias };
}

/* Checklist completo de UN médico: cada doc con subido/falta + vigencia + estado. */
function readExpedienteMedico(body) {
  try {
    var medico = (typeof body==='string') ? body : (body&&body.medico) || '';
    if (!medico) return { ok:false, error:'Falta el médico.' };
    var cfg = (readExpedienteConfig().docs||[]).filter(function(d){ return d.activo; });
    var mine = _expLeerTodos().filter(function(r){ return r.medicoKey === _expKeyMed(medico); });
    var byDoc = {}; mine.forEach(function(r){ byDoc[r.docId] = r; });
    var items = cfg.map(function(d){
      var r = byDoc[d.docId];
      var vig = d.requiereVigencia ? _expEstadoVig(r?r.fechaVigencia:'', d.diasAviso) : { estado:'no_aplica' };
      return { docId:d.docId, nombre:d.nombre, requiereVigencia:d.requiereVigencia, diasAviso:d.diasAviso,
        subido: !!(r && r.driveUrl), driveUrl:r?r.driveUrl:'', fileId:r?r.fileId:'', fechaSubida:r?r.fechaSubida:'',
        fechaVigencia:r?r.fechaVigencia:'', vigencia:vig };
    });
    var total = items.length, subidos = items.filter(function(x){ return x.subido; }).length;
    var vencidos = items.filter(function(x){ return x.vigencia.estado==='vencido'; }).length;
    var porVencer = items.filter(function(x){ return x.vigencia.estado==='por_vencer'; }).length;
    var estado = (subidos<total) ? 'incompleto' : (vencidos>0 ? 'vencido' : (porVencer>0 ? 'por_vencer' : 'completo'));
    return { ok:true, version:EXPEDIENTE_VER, medico:String(medico).trim(), items:items,
      total:total, subidos:subidos, faltan:total-subidos, vencidos:vencidos, porVencer:porVencer, estado:estado };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Carpeta de Drive del médico (bajo la raíz EXPEDIENTE_DRIVE_ID). */
function _expCarpetaMedico(medico) {
  var rootId = PropertiesService.getScriptProperties().getProperty(EXPEDIENTE_DRIVE_KEY);
  var root = rootId ? DriveApp.getFolderById(rootId) : DriveApp.createFolder('Expedientes Médicos');
  if (!rootId) PropertiesService.getScriptProperties().setProperty(EXPEDIENTE_DRIVE_KEY, root.getId());
  var name = String(medico).trim() || 'Sin nombre';
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

/* Sube un documento (base64) a Drive y lo registra (upsert por médico+docId). */
function subirDocExpediente(body) {
  try {
    body = body||{};
    if (typeof _tokenHasPermission!=='function' || !_tokenHasPermission(body.token, EXPEDIENTE_PERM))
      return { ok:false, error:'No tienes permiso para subir documentos del expediente ('+EXPEDIENTE_PERM+').' };
    var medico = String(body.medico||'').trim(); var docId = String(body.docId||'').trim();
    if (!medico || !docId) return { ok:false, error:'Falta el médico o el documento.' };
    if (!body.base64 || !body.fileName) return { ok:false, error:'Falta el archivo.' };
    var lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return { ok:false, error:'Sistema ocupado, reintenta.' };
    try {
      var carpeta = _expCarpetaMedico(medico);
      var docCfg = (readExpedienteConfig().docs||[]).filter(function(d){ return d.docId===docId; })[0] || { nombre:docId };
      var ct = body.contentType || 'application/octet-stream';
      var bytes = Utilities.base64Decode(body.base64);
      var blob = Utilities.newBlob(bytes, ct, (docCfg.nombre||docId)+' — '+body.fileName);
      // Reemplaza el archivo previo de ese doc (mismo médico+docId), si existe.
      var prev = _expLeerTodos().filter(function(r){ return r.medicoKey===_expKeyMed(medico) && r.docId===docId; })[0];
      if (prev && prev.fileId) { try { DriveApp.getFileById(prev.fileId).setTrashed(true); } catch(e){} }
      var file = carpeta.createFile(blob);
      try { file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW); } catch(e){}
      var url = file.getUrl(); var fileId = file.getId();
      var fechaVig = docCfg.requiereVigencia ? _expStr(body.fechaVigencia||'') : '';
      var sh = _expEnsure();
      if (prev) {
        sh.getRange(prev.rowNum,3,1,6).setValues([[url, fileId, _expStr(new Date()), fechaVig, _expCell(body.usuario||'sistema'), new Date()]]);
      } else {
        sh.appendRow([_expCell(medico), _expCell(docId), url, fileId, _expStr(new Date()), fechaVig, _expCell(body.usuario||'sistema'), new Date()]);
      }
      try { logAudit(body.usuario||'sistema','Expediente','Subir doc', medico+' / '+docId, '', '', body.fileName); } catch(e){}
      return { ok:true, version:EXPEDIENTE_VER, expediente: readExpedienteMedico(medico) };
    } finally { try{lock.releaseLock();}catch(e){} }
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Actualiza SOLO la fecha de vigencia de un documento ya subido. */
function actualizarVigenciaExpediente(body) {
  try {
    body = body||{};
    if (typeof _tokenHasPermission!=='function' || !_tokenHasPermission(body.token, EXPEDIENTE_PERM))
      return { ok:false, error:'No tienes permiso ('+EXPEDIENTE_PERM+').' };
    var r = _expLeerTodos().filter(function(x){ return x.medicoKey===_expKeyMed(body.medico) && x.docId===String(body.docId||'').trim(); })[0];
    if (!r) return { ok:false, error:'Ese documento aún no está subido.' };
    _expEnsure().getRange(r.rowNum,6).setValue(_expStr(body.fechaVigencia||''));
    _expEnsure().getRange(r.rowNum,8).setValue(new Date());
    try { logAudit(body.usuario||'sistema','Expediente','Vigencia', body.medico+' / '+body.docId, '', '', _expStr(body.fechaVigencia||'')); } catch(e){}
    return { ok:true, version:EXPEDIENTE_VER, expediente: readExpedienteMedico(body.medico) };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* Borra un documento del expediente (archivo de Drive a la papelera + fila). */
function eliminarDocExpediente(body) {
  try {
    body = body||{};
    if (typeof _tokenHasPermission!=='function' || !_tokenHasPermission(body.token, EXPEDIENTE_PERM))
      return { ok:false, error:'No tienes permiso ('+EXPEDIENTE_PERM+').' };
    var r = _expLeerTodos().filter(function(x){ return x.medicoKey===_expKeyMed(body.medico) && x.docId===String(body.docId||'').trim(); })[0];
    if (!r) return { ok:false, error:'Ese documento no está registrado.' };
    if (r.fileId) { try { DriveApp.getFileById(r.fileId).setTrashed(true); } catch(e){} }
    _expEnsure().deleteRow(r.rowNum);
    try { logAudit(body.usuario||'sistema','Expediente','Eliminar doc', body.medico+' / '+body.docId, '', '', ''); } catch(e){}
    return { ok:true, version:EXPEDIENTE_VER, expediente: readExpedienteMedico(body.medico) };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Médicos + su estado de expediente (para lista, badges y reporte) ────── */
function _expMedicosLista() {
  // Los médicos son los orígenes tipo 'Médico externo' (loadOrigenes → origenes.gs).
  var meds = [];
  try {
    var lo = (typeof loadOrigenes==='function') ? loadOrigenes() : null;
    var lst = (lo && lo.origenes) ? lo.origenes : [];
    lst.forEach(function(o){
      var t = (typeof _origTipoCanon==='function') ? _origTipoCanon(o.tipo) : String(o.tipo||'');
      if (t==='Médico externo' && o.activo!==false) meds.push({ nombre:o.nombre, correo:o.correo||o.email||'', telefono:o.telefono||o.tel||'' });
    });
  } catch(e){}
  return meds;
}
function readExpedientesResumen() {
  try {
    var meds = _expMedicosLista();
    var filas = meds.map(function(m){
      var e = readExpedienteMedico(m.nombre);
      return { medico:m.nombre, correo:m.correo, telefono:m.telefono, estado:e.estado||'incompleto',
        subidos:e.subidos||0, total:e.total||0, faltan:e.faltan||0, vencidos:e.vencidos||0, porVencer:e.porVencer||0,
        docsAlerta: (e.items||[]).filter(function(x){ return !x.subido || x.vigencia.estado==='vencido' || x.vigencia.estado==='por_vencer'; })
          .map(function(x){ return { nombre:x.nombre, motivo: !x.subido?'falta':x.vigencia.estado }; }) };
    });
    var accion = filas.filter(function(f){ return f.estado!=='completo'; });
    return { ok:true, version:EXPEDIENTE_VER, medicos:filas,
      resumen:{ total:filas.length, completos:filas.filter(function(f){return f.estado==='completo';}).length,
        incompletos:filas.filter(function(f){return f.estado==='incompleto';}).length,
        vencidos:filas.filter(function(f){return f.vencidos>0;}).length,
        porVencer:filas.filter(function(f){return f.porVencer>0 && f.vencidos===0;}).length },
      requierenAccion: accion.length };
  } catch(ex){ return { ok:false, error:ex.message, medicos:[] }; }
}

/* ── Notificación al médico (MANUAL o AUTOMÁTICA) ────────────────────────── */
// Arma el mensaje de estatus del expediente para un médico.
function _expMensaje(e) {
  var faltan = (e.items||[]).filter(function(x){ return !x.subido; }).map(function(x){ return x.nombre; });
  var venc   = (e.items||[]).filter(function(x){ return x.vigencia.estado==='vencido'; }).map(function(x){ return x.nombre; });
  var porv   = (e.items||[]).filter(function(x){ return x.vigencia.estado==='por_vencer'; }).map(function(x){ return x.nombre+' (vence '+x.fechaVigencia+')'; });
  var l = ['Estimado(a) '+e.medico+':', '', 'Le compartimos el estatus de su expediente médico en Hestia Fertility:'];
  if (venc.length)   l.push('', '• VENCIDOS (renovar de inmediato): '+venc.join(', '));
  if (porv.length)   l.push('', '• Por vencer: '+porv.join(', '));
  if (faltan.length) l.push('', '• Pendientes de entregar: '+faltan.join(', '));
  if (!venc.length && !porv.length && !faltan.length) l.push('', 'Su expediente está COMPLETO y al día. ¡Gracias!');
  l.push('', 'Estos documentos son requisito para realizar tratamientos en la clínica. Quedamos atentos.', '', 'Hestia Fertility · Administración');
  return l.join('\n');
}
// Envía la notificación a UN médico (correo si tiene; siempre deja recordatorio interno).
function notificarExpedienteMedico(body) {
  try {
    body = body||{};
    if (typeof _tokenHasPermission!=='function' || !_tokenHasPermission(body.token, EXPEDIENTE_PERM))
      return { ok:false, error:'No tienes permiso ('+EXPEDIENTE_PERM+').' };
    return _expNotificar(String(body.medico||'').trim(), body.usuario||'sistema', false);
  } catch(ex){ return { ok:false, error:ex.message }; }
}
function _expNotificar(medico, usuario, soloSiAlerta) {
  if (!medico) return { ok:false, error:'Falta el médico.' };
  var e = readExpedienteMedico(medico); if (!e.ok) return e;
  if (soloSiAlerta && e.estado==='completo') return { ok:true, omitido:true, medico:medico };
  var m = _expMedicosLista().filter(function(x){ return _expKeyMed(x.nombre)===_expKeyMed(medico); })[0] || {};
  var msg = _expMensaje(e); var enviado = { correo:false, recordatorio:false };
  // Correo (si el médico tiene correo y hay permiso de enviar).
  if (m.correo) {
    try { MailApp.sendEmail(m.correo, 'Expediente médico — Hestia Fertility', msg); enviado.correo = true; } catch(e2){}
  }
  // Recordatorio interno para el equipo (agenda), reusando recordatorios.gs.
  try {
    if (typeof saveRecordatorio==='function') {
      saveRecordatorio({ usuario:usuario, titulo:'Expediente: '+medico+' — '+e.estado,
        detalle:(e.faltan?e.faltan+' pendiente(s). ':'')+(e.vencidos?e.vencidos+' vencido(s). ':'')+(e.porVencer?e.porVencer+' por vencer.':''),
        fecha:_expStr(new Date()), origen:'expediente' });
      enviado.recordatorio = true;
    }
  } catch(e3){}
  try { logAudit(usuario,'Expediente','Notificar', medico, e.estado, '', (enviado.correo?'correo ':'')+(enviado.recordatorio?'recordatorio':'')); } catch(e){}
  return { ok:true, version:EXPEDIENTE_VER, medico:medico, estado:e.estado, correo:m.correo||'', enviado:enviado };
}
// Chequeo NOCTURNO: notifica automáticamente a los médicos con documentos vencidos/por vencer.
function expedienteChequeoNocturno() {
  try {
    var meds = _expMedicosLista(); var avisados = 0;
    meds.forEach(function(m){ var r = _expNotificar(m.nombre, 'sistema (nocturno)', true); if (r && r.ok && !r.omitido) avisados++; });
    return { ok:true, version:EXPEDIENTE_VER, medicos:meds.length, avisados:avisados };
  } catch(ex){ return { ok:false, error:ex.message }; }
}

/* ── Setup ───────────────────────────────────────────────────────────────── */
function setupExpediente() {
  try {
    _expEnsureCfg(); _expEnsure();
    var rootId = PropertiesService.getScriptProperties().getProperty(EXPEDIENTE_DRIVE_KEY);
    if (!rootId) { var f = DriveApp.createFolder('Expedientes Médicos'); PropertiesService.getScriptProperties().setProperty(EXPEDIENTE_DRIVE_KEY, f.getId()); rootId = f.getId(); }
    return { ok:true, version:EXPEDIENTE_VER, cfgTab:EXPEDIENTE_CFG_TAB, tab:EXPEDIENTE_TAB, driveRoot:rootId, docs:EXPEDIENTE_DEFAULTS.length };
  } catch(ex){ return { ok:false, error:ex.message }; }
}
