/**
 * nomina.gs — Módulo de NÓMINA (F1 catálogo de Empleados + F2 lectura de CFDI
 * de nómina emitidos + ISN + vacaciones). Backend Google Apps Script.
 *
 * Decisiones del usuario (2026-07-13):
 *  - Los XML de nómina están en la MISMA carpeta de emitidos que los de ingreso
 *    (onefactureXMLs/HCL2307051Y6/emitidos/{año}/{MM MES}), se filtran por
 *    TipoDeComprobante="N". Se reutiliza _facMonthFolder() de facturacion.gs.
 *  - Asimilados a salarios se timbran como CFDI de nómina (N); se distinguen por
 *    el TipoRegimen del receptor de nómina (02 = sueldos; 05–11 = asimilados).
 *  - ISN 3% por defecto, CONFIGURABLE (Script Property NOMINA_CFG).
 *  - Al validar → una orden de pago por empleado a Cuentas por Pagar (fase F4).
 *
 * Persistencia:
 *  - Empleados: hoja "Empleados" en SHEET_ID (junto a Usuarios/Roles).
 *  - Config: Script Property NOMINA_CFG (JSON).
 *  - La nómina del mes se LEE en vivo de los XML (no se duplica en hoja); lo que
 *    se captura (bonos, validación, días de vacaciones tomados) irá en hojas
 *    aparte en las fases F3/F4.
 *
 * Rutas (se agregan en core.gs GET / finance.gs POST):
 *   GET  empleados                         → readEmpleados()
 *   GET  nominaMes&anio=&mes=              → readNominaMes(anio, mes)
 *   GET  nominaCfg                          → nominaCfgRead()
 *   GET  misRecibos&anio=&mes=            → nominaMisRecibos(token, anio, mes)
 *   POST saveEmpleado {token, emp}          → saveEmpleado(body)
 *   POST deleteEmpleado {token, numEmpleado}→ deleteEmpleado(body)
 *   POST nominaCfgSave {token, cfg}         → nominaCfgSave(body)
 */

// ── Configuración ──────────────────────────────────────────────────────
var NOMINA_CFG_KEY = 'NOMINA_CFG';
function _nominaCfgDefault() {
  return {
    isnTasa: 3,            // % Impuesto Sobre Nómina (configurable por estado)
    isnBase: 'percepciones', // 'percepciones' (total) | 'gravado'
    cxpProveedor: 'Nómina',  // proveedor/beneficiario en Cuentas por Pagar
    cxpCuentaContable: 'Sueldos y salarios'
  };
}
function _nominaCfg() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(NOMINA_CFG_KEY);
    if (!raw) return _nominaCfgDefault();
    var o = JSON.parse(raw), d = _nominaCfgDefault();
    for (var k in d) { if (o[k] === undefined) o[k] = d[k]; }
    return o;
  } catch (e) { return _nominaCfgDefault(); }
}
function nominaCfgRead() { return { ok: true, cfg: _nominaCfg() }; }
function nominaCfgSave(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) {
      return { ok: false, error: 'Sin autorización para configurar nómina (editar_egresos).' };
    }
    var cur = _nominaCfg(), inc = body.cfg || {};
    for (var k in inc) { cur[k] = inc[k]; }
    if (inc.isnTasa !== undefined) { var t = parseFloat(inc.isnTasa); cur.isnTasa = isNaN(t) ? 3 : t; }
    PropertiesService.getScriptProperties().setProperty(NOMINA_CFG_KEY, JSON.stringify(cur));
    return { ok: true, cfg: cur };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── Catálogo de Empleados (hoja "Empleados" en SHEET_ID) ────────────────
var NOM_EMP_TAB = 'Empleados';
var NOM_EMP_HEADERS = [
  'NumEmpleado', 'Nombre', 'RFC', 'CURP', 'NSS', 'Puesto', 'Departamento',
  'FechaIngreso', 'SalarioDiario', 'SBC', 'Periodicidad', 'Tipo',
  'Banco', 'CLABE', 'UsuarioEmail', 'Activo', 'Notas'
];
function _nomEmpSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(NOM_EMP_TAB);
  if (!sh) {
    sh = ss.insertSheet(NOM_EMP_TAB);
    sh.getRange(1, 1, 1, NOM_EMP_HEADERS.length).setValues([NOM_EMP_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // Append de columnas nuevas (por nombre; no reordena datos existentes).
  var last = Math.max(1, sh.getLastColumn());
  var ex = sh.getRange(1, 1, 1, last).getValues()[0].map(function (v) { return String(v); });
  var add = [];
  for (var i = 0; i < NOM_EMP_HEADERS.length; i++) if (ex.indexOf(NOM_EMP_HEADERS[i]) === -1) add.push(NOM_EMP_HEADERS[i]);
  if (add.length) { sh.getRange(1, ex.length + 1, 1, add.length).setValues([add]); sh.getRange(1, 1, 1, ex.length + add.length).setFontWeight('bold'); }
  return sh;
}
function readEmpleados() {
  try {
    var sh = _nomEmpSheet();
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, empleados: [] };
    var h = data[0], out = [];
    for (var i = 1; i < data.length; i++) {
      var o = {};
      for (var c = 0; c < h.length; c++) o[String(h[c])] = data[i][c];
      if (!o.NumEmpleado && !o.Nombre) continue;
      if (o.FechaIngreso instanceof Date) o.FechaIngreso = Utilities.formatDate(o.FechaIngreso, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd');
      o.Activo = (String(o.Activo).toLowerCase() !== 'no' && String(o.Activo).toLowerCase() !== 'false' && o.Activo !== false);
      o.diasVacaciones = _nomDiasVacaciones(_nomAntiguedadAnios(o.FechaIngreso));
      out.push(o);
    }
    out.sort(function (a, b) { return String(a.Nombre || '').localeCompare(String(b.Nombre || '')); });
    return { ok: true, empleados: out };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
function saveEmpleado(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) {
      return { ok: false, error: 'Sin autorización para editar empleados (editar_egresos).' };
    }
    var e = body.emp || {};
    var num = String(e.NumEmpleado || '').trim();
    if (!num) return { ok: false, error: 'El número de empleado es obligatorio.' };
    var sh = _nomEmpSheet();
    var data = sh.getDataRange().getValues();
    var h = data[0];
    function idx(n) { return h.indexOf(n); }
    var numCols = { SalarioDiario: 1, SBC: 1 };
    var vals = h.map(function (col) {
      col = String(col);
      if (col === 'NumEmpleado') return num;
      if (col === 'Activo') return (e.Activo === false || String(e.Activo).toLowerCase() === 'no') ? 'No' : 'Sí';
      var v = e[col];
      if (v === undefined || v === null) return numCols[col] ? 0 : '';
      if (numCols[col]) { var n = parseFloat(String(v).replace(/[$,\s]/g, '')); return isNaN(n) ? 0 : n; }
      return v;
    });
    var found = -1, iNum = idx('NumEmpleado');
    for (var i = 1; i < data.length; i++) { if (String(data[i][iNum] || '').trim() === num) { found = i + 1; break; } }
    if (found > 0) sh.getRange(found, 1, 1, h.length).setValues([vals]);
    else sh.appendRow(vals);
    try { logAudit((verifyToken(body.token || '') || 'sistema'), 'Nómina', found > 0 ? 'EditarEmpleado' : 'CrearEmpleado', num + ' ' + (e.Nombre || ''), '', '', ''); } catch (x) {}
    return { ok: true, numEmpleado: num };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
function deleteEmpleado(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización.' };
    var num = String(body.numEmpleado || '').trim();
    if (!num) return { ok: false, error: 'numEmpleado requerido' };
    var sh = _nomEmpSheet();
    var data = sh.getDataRange().getValues();
    var iNum = data[0].indexOf('NumEmpleado');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][iNum] || '').trim() === num) { sh.deleteRow(i + 1); return { ok: true }; }
    }
    return { ok: false, error: 'Empleado no encontrado' };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── Vacaciones (LFT reforma 2023) y antigüedad ─────────────────────────
function _nomAntiguedadAnios(fechaIngreso) {
  try {
    if (!fechaIngreso) return 0;
    var d = new Date(String(fechaIngreso).substring(0, 10) + 'T00:00:00');
    if (isNaN(d.getTime())) return 0;
    var hoy = new Date();
    var anios = (hoy.getTime() - d.getTime()) / (365.25 * 24 * 3600 * 1000);
    return anios < 0 ? 0 : anios;
  } catch (e) { return 0; }
}
function _nomDiasVacaciones(anios) {
  anios = Math.floor(anios || 0);
  if (anios < 1) return 0;
  var tabla = [0, 12, 14, 16, 18, 20]; // años 1..5
  if (anios <= 5) return tabla[anios];
  return 22 + Math.floor((anios - 6) / 5) * 2; // 6–10=22, 11–15=24, 16–20=26 …
}

// ── Lectura de CFDI de nómina (complemento nomina12) ───────────────────
function _nomAttr(name, scope) {
  var m = String(scope || '').match(new RegExp(name + '\\s*=\\s*"([^"]*)"'));
  return m ? m[1] : '';
}
// Régimen de asimilados a salarios (05–11 en el catálogo SAT). 02 = sueldos.
function _nomEsAsimilado(tipoRegimen) {
  return ['05', '06', '07', '08', '09', '10', '11'].indexOf(String(tipoRegimen)) > -1;
}
function _nomParseCfdi(xml) {
  var compM = xml.match(/<(\w+):Comprobante\b[^>]*>/);
  var comp = compM ? compM[0] : xml.substring(0, 3000);
  var tipo = _nomAttr('TipoDeComprobante', comp);
  if (tipo !== 'N') return { tipo: tipo };
  var total = parseFloat(_nomAttr('Total', comp)) || 0;
  var fecha = _nomAttr('Fecha', comp);
  var recM = xml.match(/<(\w+):Receptor\b[^>]*>/); // primero el receptor fiscal (cfdi)
  var rec = recM ? recM[0] : '';
  var rfc = _nomAttr('Rfc', rec), nombre = _nomAttr('Nombre', rec);
  var uuidM = xml.match(/UUID\s*=\s*"([^"]*)"/i);
  var uuid = uuidM ? uuidM[1] : '';

  // Bloque de nómina completo
  var nbM = xml.match(/<(\w+):Nomina\b[\s\S]*?<\/\1:Nomina>/);
  var nb = nbM ? nbM[0] : '';
  var nomOpenM = nb.match(/<(\w+):Nomina\b[^>]*>/);
  var nomOpen = nomOpenM ? nomOpenM[0] : '';
  var tipoNomina = _nomAttr('TipoNomina', nomOpen);
  var totPerc = parseFloat(_nomAttr('TotalPercepciones', nomOpen)) || 0;
  var totDed = parseFloat(_nomAttr('TotalDeducciones', nomOpen)) || 0;
  var totOtros = parseFloat(_nomAttr('TotalOtrosPagos', nomOpen)) || 0;
  var numDias = parseFloat(_nomAttr('NumDiasPagados', nomOpen)) || 0;
  var fechaPago = _nomAttr('FechaPago', nomOpen);
  var fechaIni = _nomAttr('FechaInicialPago', nomOpen);
  var fechaFin = _nomAttr('FechaFinalPago', nomOpen);

  // Receptor de nómina (Curp/NumEmpleado/Puesto/TipoRegimen)
  var nrM = nb.match(/<(\w+):Receptor\b[^>]*>/);
  var nr = nrM ? nrM[0] : '';
  var curp = _nomAttr('Curp', nr), numEmp = _nomAttr('NumEmpleado', nr);
  var depto = _nomAttr('Departamento', nr), puesto = _nomAttr('Puesto', nr);
  var tipoRegimen = _nomAttr('TipoRegimen', nr);
  var periodicidad = _nomAttr('PeriodicidadPago', nr);
  var fechaIngreso = _nomAttr('FechaInicioRelLaboral', nr);

  // Percepciones (totales + prima vacacional 021 / aguinaldo 002)
  var pnM = nb.match(/<(\w+):Percepciones\b[^>]*>/);
  var pn = pnM ? pnM[0] : '';
  var totSueldos = parseFloat(_nomAttr('TotalSueldos', pn)) || 0;
  var totGravado = parseFloat(_nomAttr('TotalGravado', pn)) || 0;
  var totExento = parseFloat(_nomAttr('TotalExento', pn)) || 0;
  var primaVac = 0, aguinaldo = 0;
  var percRe = /<(\w+):Percepcion\b[^>]*>/g, pm;
  while ((pm = percRe.exec(nb))) {
    var seg = pm[0], tp = _nomAttr('TipoPercepcion', seg);
    var imp = (parseFloat(_nomAttr('ImporteGravado', seg)) || 0) + (parseFloat(_nomAttr('ImporteExento', seg)) || 0);
    if (tp === '021') primaVac += imp;
    if (tp === '002') aguinaldo += imp;
  }
  // Deducciones (ISR 002, IMSS 001)
  var isr = 0, imss = 0;
  var dedRe = /<(\w+):Deduccion\b[^>]*>/g, dm;
  while ((dm = dedRe.exec(nb))) {
    var ds = dm[0], td = _nomAttr('TipoDeduccion', ds), impo = parseFloat(_nomAttr('Importe', ds)) || 0;
    if (td === '002') isr += impo;
    if (td === '001') imss += impo;
  }

  var neto = totPerc + totOtros - totDed;
  return {
    tipo: 'N', uuid: uuid, total: total, fecha: (fecha || '').substring(0, 10),
    rfc: rfc, nombre: nombre, curp: curp, numEmpleado: numEmp,
    departamento: depto, puesto: puesto, tipoRegimen: tipoRegimen,
    asimilado: _nomEsAsimilado(tipoRegimen), tipoNomina: tipoNomina,
    periodicidad: periodicidad, fechaIngreso: fechaIngreso,
    fechaPago: fechaPago, fechaInicial: fechaIni, fechaFinal: fechaFin, numDias: numDias,
    totalPercepciones: totPerc, totalDeducciones: totDed, totalOtrosPagos: totOtros,
    sueldos: totSueldos, gravado: totGravado, exento: totExento,
    primaVacacional: primaVac, aguinaldo: aguinaldo, isrRetenido: isr, imss: imss, neto: neto
  };
}

// Lee todos los CFDI de nómina de un mes, agrupados por empleado. Calcula ISN.
function readNominaMes(anio, mes) {
  try {
    anio = parseInt(anio, 10); mes = parseInt(mes, 10);
    if (!anio || !mes) return { ok: false, error: 'anio y mes requeridos' };
    if (typeof _facMonthFolder !== 'function') return { ok: false, error: 'facturacion.gs (carpetas de emitidos) no está disponible.' };
    var folder = _facMonthFolder(anio, mes);
    var cfg = _nominaCfg();
    if (!folder) {
      return { ok: true, anio: anio, mes: mes, encontrada: false, empleados: [], recibos: [],
        totales: { neto: 0, percepciones: 0, isnBase: 0, isnTasa: cfg.isnTasa, isn: 0, numRecibos: 0, numEmpleados: 0 } };
    }
    var files = folder.getFiles();
    var recibos = [], porEmp = {};
    while (files.hasNext()) {
      var f = files.next();
      var name = f.getName();
      if (!/\.xml$/i.test(name)) continue;
      var xml; try { xml = f.getBlob().getDataAsString('UTF-8'); } catch (e) { continue; }
      if (xml.indexOf(':Nomina') === -1 && xml.indexOf('TipoDeComprobante="N"') === -1) continue;
      var p = _nomParseCfdi(xml);
      if (p.tipo !== 'N') continue;
      p.fileId = f.getId(); p.fileUrl = f.getUrl(); p.fileName = name;
      recibos.push(p);
      var key = p.numEmpleado || p.rfc || p.curp || p.nombre || ('r' + recibos.length);
      if (!porEmp[key]) porEmp[key] = {
        clave: key, numEmpleado: p.numEmpleado, rfc: p.rfc, nombre: p.nombre, curp: p.curp,
        departamento: p.departamento, puesto: p.puesto, asimilado: p.asimilado, tipoRegimen: p.tipoRegimen,
        recibos: 0, totalPercepciones: 0, totalDeducciones: 0, totalOtrosPagos: 0,
        gravado: 0, primaVacacional: 0, aguinaldo: 0, isrRetenido: 0, imss: 0, neto: 0, detalle: []
      };
      var e = porEmp[key];
      e.recibos++; e.totalPercepciones += p.totalPercepciones; e.totalDeducciones += p.totalDeducciones;
      e.totalOtrosPagos += p.totalOtrosPagos; e.gravado += p.gravado; e.primaVacacional += p.primaVacacional;
      e.aguinaldo += p.aguinaldo; e.isrRetenido += p.isrRetenido; e.imss += p.imss; e.neto += p.neto;
      e.detalle.push({ uuid: p.uuid, fecha: p.fecha, fechaPago: p.fechaPago, tipoNomina: p.tipoNomina,
        totalPercepciones: p.totalPercepciones, totalDeducciones: p.totalDeducciones, totalOtrosPagos: p.totalOtrosPagos,
        neto: p.neto, fileUrl: p.fileUrl, fileId: p.fileId });
    }
    var empleados = [];
    for (var k in porEmp) if (porEmp.hasOwnProperty(k)) empleados.push(porEmp[k]);
    empleados.sort(function (a, b) { return String(a.nombre || '').localeCompare(String(b.nombre || '')); });

    var baseISN = 0, totNeto = 0, totPercT = 0;
    empleados.forEach(function (e) {
      baseISN += (cfg.isnBase === 'gravado' ? e.gravado : e.totalPercepciones);
      totNeto += e.neto; totPercT += e.totalPercepciones;
    });
    var isn = baseISN * ((parseFloat(cfg.isnTasa) || 0) / 100);
    return {
      ok: true, anio: anio, mes: mes, encontrada: true, empleados: empleados, recibos: recibos,
      totales: { neto: totNeto, percepciones: totPercT, isnBase: baseISN, isnTasa: cfg.isnTasa, isn: isn,
        numRecibos: recibos.length, numEmpleados: empleados.length }
    };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── Autoservicio: recibos del empleado dueño del token ─────────────────
function nominaMisRecibos(token, anio, mes) {
  try {
    var email = ''; try { email = verifyToken(token || '') || ''; } catch (e) {}
    if (!email) return { ok: false, error: 'Sesión inválida.' };
    // Buscar el empleado vinculado a este usuario.
    var emps = readEmpleados().empleados || [];
    var yo = null;
    for (var i = 0; i < emps.length; i++) {
      if (String(emps[i].UsuarioEmail || '').trim().toLowerCase() === email.toLowerCase()) { yo = emps[i]; break; }
    }
    if (!yo) return { ok: true, vinculado: false, recibos: [] };
    var data = readNominaMes(anio, mes);
    if (!data.ok) return data;
    var miRfc = String(yo.RFC || '').toUpperCase(), miNum = String(yo.NumEmpleado || '');
    var mios = (data.recibos || []).filter(function (r) {
      return (miRfc && String(r.rfc || '').toUpperCase() === miRfc) || (miNum && String(r.numEmpleado || '') === miNum);
    });
    return { ok: true, vinculado: true, empleado: { NumEmpleado: yo.NumEmpleado, Nombre: yo.Nombre, Puesto: yo.Puesto }, recibos: mios };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── Diagnóstico: correr desde el editor (selecciona nominaDiagnostico → Run) ──
// Cuenta los XML de la carpeta de emitidos de un mes, los clasifica por tipo, y
// corre readNominaMes IMPRIMIENDO el resultado en el registro de ejecución.
// Cambia ANIO/MES abajo si quieres otro periodo.
function nominaDiagnostico() {
  var ANIO = 2026, MES = 6; // ← ajusta aquí el mes a validar
  var out = ['== Diagnóstico Nómina ' + ANIO + '-' + (MES < 10 ? '0' : '') + MES + ' =='];
  try {
    if (typeof _facMonthFolder !== 'function') { out.push('ERROR: facturacion.gs no está en el proyecto (no encuentro _facMonthFolder).'); Logger.log(out.join('\n')); return; }
    var mesTag = (MES < 10 ? '0' : '') + MES + ' ' + FAC_MESES_ABR[MES - 1];
    var folder = _facMonthFolder(ANIO, MES);
    if (!folder) {
      out.push('❌ Carpeta NO encontrada: onefactureXMLs/HCL2307051Y6/emitidos/' + ANIO + '/' + mesTag);
      out.push('   Revisa que los nombres de carpeta coincidan (¿la nómina se timbra ahí?).');
      Logger.log(out.join('\n')); return;
    }
    out.push('✔ Carpeta: ' + folder.getName() + '  ' + folder.getUrl());
    var files = folder.getFiles(), totXml = 0, tipos = {};
    while (files.hasNext()) {
      var f = files.next(); if (!/\.xml$/i.test(f.getName())) continue; totXml++;
      var xml; try { xml = f.getBlob().getDataAsString('UTF-8'); } catch (e) { continue; }
      var m = xml.match(/TipoDeComprobante\s*=\s*"?([A-Z])"?/);
      var t = m ? m[1] : '?'; tipos[t] = (tipos[t] || 0) + 1;
    }
    out.push('XML en la carpeta: ' + totXml + '  · por tipo: ' + JSON.stringify(tipos) + '  (N = nómina)');
    var r = readNominaMes(ANIO, MES);
    out.push('readNominaMes.ok = ' + r.ok + (r.error ? ('  error=' + r.error) : ''));
    if (r.ok) {
      out.push('Recibos de nómina: ' + r.totales.numRecibos + '  ·  Empleados: ' + r.totales.numEmpleados);
      out.push('Neto total: $' + (r.totales.neto || 0).toFixed(2) + '  ·  ISN (' + r.totales.isnTasa + '%): $' + (r.totales.isn || 0).toFixed(2));
      (r.empleados || []).forEach(function (e) {
        out.push('  • ' + (e.nombre || e.rfc || e.numEmpleado || '¿?') + '  | recibos=' + e.recibos +
          ' | neto=$' + (e.neto || 0).toFixed(2) + ' | ISR=$' + (e.isrRetenido || 0).toFixed(2) +
          (e.primaVacacional ? ' | primaVac=$' + e.primaVacacional.toFixed(2) : '') + (e.asimilado ? ' | ASIMILADO' : ''));
      });
    }
  } catch (ex) { out.push('EXCEPCIÓN: ' + ex.message); }
  Logger.log(out.join('\n'));
}
