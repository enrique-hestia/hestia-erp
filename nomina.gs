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
    isnTasa: 4,            // % ISN — CDMX: 4% desde 2025-01-01 (readNominaMes lo ajusta por periodo)
    isnBase: 'percepciones', // 'percepciones' (total) | 'gravado'
    isnExcluirAsimilados: true, // CDMX: el ISN aplica a sueldos/salarios, NO a asimilados
    isnSubsidio: false,     // subsidio CDMX para empresas pequeñas (baja la tasa)
    isnSubsidioPct: 1,      // puntos de reducción: 1 (micro ≤10) | 0.5 (pequeña 11–50)
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
  var primaVac = 0, aguinaldo = 0, percDet = [];
  var percRe = /<(\w+):Percepcion\b[^>]*>/g, pm;
  while ((pm = percRe.exec(nb))) {
    var seg = pm[0], tp = _nomAttr('TipoPercepcion', seg);
    var g = parseFloat(_nomAttr('ImporteGravado', seg)) || 0, ex = parseFloat(_nomAttr('ImporteExento', seg)) || 0;
    var imp = g + ex;
    percDet.push({ tipo: tp, clave: _nomAttr('Clave', seg), concepto: _nomAttr('Concepto', seg), gravado: g, exento: ex, importe: imp });
    if (tp === '021') primaVac += imp;
    if (tp === '002') aguinaldo += imp;
  }
  // Deducciones (ISR 002, IMSS 001)
  var isr = 0, imss = 0, dedDet = [];
  var dedRe = /<(\w+):Deduccion\b[^>]*>/g, dm;
  while ((dm = dedRe.exec(nb))) {
    var ds = dm[0], td = _nomAttr('TipoDeduccion', ds), impo = parseFloat(_nomAttr('Importe', ds)) || 0;
    dedDet.push({ tipo: td, clave: _nomAttr('Clave', ds), concepto: _nomAttr('Concepto', ds), importe: impo });
    if (td === '002') isr += impo;
    if (td === '001') imss += impo;
  }
  // Otros pagos (subsidio al empleo, etc.)
  var otrosDet = [];
  var opRe = /<(\w+):OtroPago\b[^>]*>/g, om;
  while ((om = opRe.exec(nb))) {
    var os = om[0];
    otrosDet.push({ tipo: _nomAttr('TipoOtroPago', os), clave: _nomAttr('Clave', os), concepto: _nomAttr('Concepto', os), importe: parseFloat(_nomAttr('Importe', os)) || 0 });
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
    primaVacacional: primaVac, aguinaldo: aguinaldo, isrRetenido: isr, imss: imss, neto: neto,
    percepciones: percDet, deducciones: dedDet, otrosPagos: otrosDet
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
        neto: p.neto, fileUrl: p.fileUrl, fileId: p.fileId,
        // Detalle completo para el visor de recibo:
        nombre: p.nombre, rfc: p.rfc, curp: p.curp, numEmpleado: p.numEmpleado, puesto: p.puesto, departamento: p.departamento,
        periodicidad: p.periodicidad, fechaInicial: p.fechaInicial, fechaFinal: p.fechaFinal, numDias: p.numDias,
        percepciones: p.percepciones, deducciones: p.deducciones, otrosPagos: p.otrosPagos });
    }
    var empleados = [];
    for (var k in porEmp) if (porEmp.hasOwnProperty(k)) empleados.push(porEmp[k]);
    empleados.sort(function (a, b) { return String(a.nombre || '').localeCompare(String(b.nombre || '')); });

    // ISN: en CDMX aplica SOLO a sueldos/salarios (no a asimilados). Se separan
    // los totales por tipo y la base del ISN excluye asimilados (configurable).
    var baseISN = 0, totNeto = 0, totPercT = 0;
    var netoSueldos = 0, netoAsim = 0, numSueldos = 0, numAsim = 0, percSueldos = 0, percAsim = 0;
    empleados.forEach(function (e) {
      var esAsim = !!e.asimilado;
      var baseE = (cfg.isnBase === 'gravado' ? e.gravado : e.totalPercepciones);
      if (!esAsim || !cfg.isnExcluirAsimilados) baseISN += baseE;
      totNeto += e.neto; totPercT += e.totalPercepciones;
      if (esAsim) { netoAsim += e.neto; numAsim++; percAsim += e.totalPercepciones; }
      else { netoSueldos += e.neto; numSueldos++; percSueldos += e.totalPercepciones; }
    });
    // CDMX: la tasa del ISN depende del periodo (3% hasta 2024, 4% desde 2025-01-01).
    var isnTasaAplicada = (anio >= 2025) ? 4 : 3;
    // Subsidio CDMX para empresas pequeñas (solo cuando la tasa base es 4%): micro
    // ≤10 empleados −1% (→3%), pequeña 11–50 −0.5% (→3.5%). Ya viene restado en isn.
    var subsPct = (cfg.isnSubsidio && anio >= 2025) ? (parseFloat(cfg.isnSubsidioPct) || 0) : 0;
    var isnTasaEfectiva = isnTasaAplicada - subsPct;
    var isnSubsidioMonto = baseISN * (subsPct / 100);
    var isn = baseISN * (isnTasaEfectiva / 100);
    return {
      ok: true, anio: anio, mes: mes, encontrada: true, empleados: empleados, recibos: recibos,
      totales: {
        neto: totNeto, percepciones: totPercT, isnBase: baseISN, isnTasa: isnTasaAplicada,
        isnTasaEfectiva: isnTasaEfectiva, isnSubsidio: (!!cfg.isnSubsidio && anio >= 2025),
        isnSubsidioPct: subsPct, isnSubsidioMonto: isnSubsidioMonto, isn: isn,
        isnExcluirAsimilados: !!cfg.isnExcluirAsimilados,
        numRecibos: recibos.length, numEmpleados: empleados.length,
        netoSueldos: netoSueldos, netoAsimilados: netoAsim, numSueldos: numSueldos, numAsimilados: numAsim,
        percepcionesSueldos: percSueldos, percepcionesAsimilados: percAsim
      }
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

/* ═══════════════════════════════════════════════════════════════════════
 * F3/F4 — Vínculo Usuario↔Empleado · Bonos por empleado/mes · Validación
 * del mes a Cuentas por Pagar (una orden por empleado).
 *
 * Hojas nuevas (se crean solas al primer uso):
 *   - "Nomina_Bonos" (en SHEET_ID): bonos capturados por empleado/mes.
 *   - "Nomina_Meses" (en SHEET_ID): estado Borrador/Validada por periodo.
 * Las órdenes de pago se escriben en Egresos2026 (mismo mecanismo que Gastos
 * Fijos: fila sin fecha de pago en col B + Pagado=false en col N ⇒ aparece
 * como Cuenta por Pagar en la vista CxP, que lee readBDCxP de Egresos2026).
 *
 * Rutas nuevas (registrar en core.gs GET / finance.gs POST):
 *   GET  nominaMesEstado&anio=&mes=          → nominaMesEstado(anio, mes)
 *   GET  nominaBonos&anio=&mes=              → readBonos(anio, mes)
 *   POST vincularEmpleadoUsuario {token, email, numEmpleado, nombre, vincular}
 *   POST nominaValidarMes {token, anio, mes, forzar}
 *   POST saveBono {token, anio, mes, numEmpleado, nombre, concepto, monto}
 *   POST deleteBono {token, bonoId}
 * ═══════════════════════════════════════════════════════════════════════ */

// ── Utilidades de periodo/fecha ────────────────────────────────────────
function _nomPeriodo(anio, mes) {
  var a = parseInt(anio, 10), m = parseInt(mes, 10);
  if (!a || !m || m < 1 || m > 12) return '';
  return a + '-' + (m < 10 ? '0' : '') + m;
}
function _nomFinDeMes(anio, mes) {
  var a = parseInt(anio, 10), m = parseInt(mes, 10);
  var last = new Date(a, m, 0).getDate(); // día 0 del mes siguiente = último día
  return a + '-' + (m < 10 ? '0' : '') + m + '-' + (last < 10 ? '0' : '') + last;
}
function _nomMesNombreEs(mes) {
  var M = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  return M[(parseInt(mes, 10) || 1) - 1] || String(mes);
}
function _nomFechaStr(v) {
  try {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd');
    return String(v).substring(0, 10);
  } catch (e) { return String(v || ''); }
}

// ── A. Vínculo Usuario ↔ Empleado ──────────────────────────────────────
// Merge no destructivo sobre la hoja Empleados: conserva el resto de la fila
// si el empleado ya existía; sólo fija NumEmpleado/Nombre/UsuarioEmail/Activo.
// vincular:false ⇒ desvincula (limpia UsuarioEmail) sin borrar al empleado.
function vincularEmpleadoUsuario(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) {
      return { ok: false, error: 'Sin autorización para vincular empleados (editar_egresos).' };
    }
    var email = String(body.email || '').trim().toLowerCase();
    if (!email) return { ok: false, error: 'Email del usuario requerido.' };
    var vincular = (body.vincular === undefined) ? true : !!body.vincular;
    var num = String(body.numEmpleado || '').trim();
    var nombre = String(body.nombre || '').trim();
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}

    var sh = _nomEmpSheet();
    var data = sh.getDataRange().getValues();
    var h = data[0];
    var iNum = h.indexOf('NumEmpleado'), iNom = h.indexOf('Nombre'),
        iEmail = h.indexOf('UsuarioEmail'), iActivo = h.indexOf('Activo');
    if (iEmail < 0) return { ok: false, error: 'La hoja Empleados no tiene columna UsuarioEmail.' };

    // ── Desvincular: limpia UsuarioEmail de cualquier empleado con este email ──
    if (!vincular) {
      var cleared = 0;
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][iEmail] || '').trim().toLowerCase() === email) {
          sh.getRange(r + 1, iEmail + 1).setValue(''); cleared++;
        }
      }
      try { logAudit(usuario || 'sistema', 'Nómina', 'DesvincularEmpleado', email, '', '', String(cleared)); } catch (x) {}
      return { ok: true, desvinculado: true, filas: cleared };
    }

    if (!num) return { ok: false, error: 'El número de empleado es obligatorio para vincular.' };

    // 1 usuario ↔ 1 empleado: si OTRO empleado tenía este email, límpialo.
    for (var r2 = 1; r2 < data.length; r2++) {
      if (String(data[r2][iEmail] || '').trim().toLowerCase() === email &&
          String(data[r2][iNum] || '').trim() !== num) {
        sh.getRange(r2 + 1, iEmail + 1).setValue('');
      }
    }
    // Buscar la fila del empleado; si existe → merge; si no → crear con lo mínimo.
    var found = -1;
    for (var r3 = 1; r3 < data.length; r3++) {
      if (String(data[r3][iNum] || '').trim() === num) { found = r3; break; }
    }
    if (found > -1) {
      var rowNum = found + 1;
      sh.getRange(rowNum, iEmail + 1).setValue(email);
      if (nombre && iNom > -1 && !String(data[found][iNom] || '').trim()) sh.getRange(rowNum, iNom + 1).setValue(nombre);
      if (iActivo > -1) sh.getRange(rowNum, iActivo + 1).setValue('Sí');
      try { logAudit(usuario || 'sistema', 'Nómina', 'VincularEmpleado', num + ' ' + email, '', '', 'merge'); } catch (x) {}
      return { ok: true, numEmpleado: num, creado: false };
    }
    var vals = h.map(function (col) {
      col = String(col);
      if (col === 'NumEmpleado') return num;
      if (col === 'Nombre') return nombre;
      if (col === 'UsuarioEmail') return email;
      if (col === 'Activo') return 'Sí';
      if (col === 'SalarioDiario' || col === 'SBC') return 0;
      return '';
    });
    sh.appendRow(vals);
    try { logAudit(usuario || 'sistema', 'Nómina', 'VincularEmpleado', num + ' ' + email, '', '', 'nuevo'); } catch (x) {}
    return { ok: true, numEmpleado: num, creado: true };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── C. Bonos por empleado/mes (hoja Nomina_Bonos) ──────────────────────
var NOM_BONOS_TAB = 'Nomina_Bonos';
var NOM_BONOS_HEADERS = ['BonoID', 'Periodo', 'NumEmpleado', 'Nombre', 'Concepto', 'Monto', 'Usuario', 'Fecha'];
function _nomBonosSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(NOM_BONOS_TAB);
  if (!sh) {
    sh = ss.insertSheet(NOM_BONOS_TAB);
    sh.getRange(1, 1, 1, NOM_BONOS_HEADERS.length).setValues([NOM_BONOS_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function readBonos(anio, mes) {
  try {
    var per = _nomPeriodo(anio, mes);
    if (!per) return { ok: false, error: 'anio y mes requeridos', bonos: [] };
    var sh = _nomBonosSheet();
    var data = sh.getDataRange().getValues();
    var out = [], total = 0;
    for (var i = 1; i < data.length; i++) {
      var t = data[i];
      if (String(t[1] || '') !== per) continue;
      var monto = parseFloat(String(t[5] || '').replace(/[$,\s]/g, '')) || 0;
      out.push({ bonoId: String(t[0] || ''), periodo: per, numEmpleado: String(t[2] || ''), nombre: String(t[3] || ''),
        concepto: String(t[4] || ''), monto: monto, usuario: String(t[6] || ''), fecha: _nomFechaStr(t[7]) });
      total += monto;
    }
    return { ok: true, periodo: per, bonos: out, total: total };
  } catch (ex) { return { ok: false, error: ex.message, bonos: [] }; }
}
function saveBono(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var per = _nomPeriodo(body.anio, body.mes);
    if (!per) return { ok: false, error: 'anio y mes requeridos' };
    var num = String(body.numEmpleado || '').trim();
    if (!num) return { ok: false, error: 'numEmpleado requerido' };
    var concepto = String(body.concepto || 'Bono').trim() || 'Bono';
    var monto = parseFloat(String(body.monto || '').replace(/[$,\s]/g, '')) || 0;
    if (!monto) return { ok: false, error: 'El monto del bono debe ser mayor a cero.' };
    var sh = _nomBonosSheet();
    var id = 'BON-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
    sh.appendRow([id, per, num, String(body.nombre || ''), concepto, monto, usuario, new Date()]);
    try { logAudit(usuario || 'sistema', 'Nómina', 'CrearBono', num + ' ' + per, '', '', concepto + ' $' + monto); } catch (x) {}
    return { ok: true, bonoId: id };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
function deleteBono(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var id = String(body.bonoId || '').trim();
    if (!id) return { ok: false, error: 'bonoId requerido' };
    var sh = _nomBonosSheet();
    var data = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || '').trim() === id) {
        sh.deleteRow(i + 1);
        try { logAudit(verifyToken(body.token || '') || 'sistema', 'Nómina', 'BorrarBono', id, '', '', ''); } catch (x) {}
        return { ok: true };
      }
    }
    return { ok: false, error: 'Bono no encontrado' };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
// Suma de bonos del mes por NumEmpleado. {porNum:{num->total}, total, list}.
function _nomBonosDelMes(periodo) {
  var sh = _nomBonosSheet();
  var data = sh.getDataRange().getValues();
  var porNum = {}, total = 0, list = [];
  for (var i = 1; i < data.length; i++) {
    var t = data[i];
    if (String(t[1] || '') !== periodo) continue;
    var num = String(t[2] || '').trim();
    var monto = parseFloat(String(t[5] || '').replace(/[$,\s]/g, '')) || 0;
    porNum[num] = (porNum[num] || 0) + monto;
    total += monto;
    list.push({ numEmpleado: num, concepto: String(t[4] || ''), monto: monto });
  }
  return { porNum: porNum, total: total, list: list };
}

// ── B. Estado del mes (Nomina_Meses) + validación a Cuentas por Pagar ───
var NOM_MESES_TAB = 'Nomina_Meses';
var NOM_MESES_HEADERS = ['Periodo', 'Estado', 'FechaValidacion', 'Usuario', 'NumOrdenes', 'TotalNeto', 'Notas'];
function _nomMesesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(NOM_MESES_TAB);
  if (!sh) {
    sh = ss.insertSheet(NOM_MESES_TAB);
    sh.getRange(1, 1, 1, NOM_MESES_HEADERS.length).setValues([NOM_MESES_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function nominaMesEstado(anio, mes) {
  try {
    var per = _nomPeriodo(anio, mes);
    if (!per) return { ok: false, error: 'anio y mes requeridos' };
    var sh = _nomMesesSheet();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '') === per) {
        return { ok: true, periodo: per, encontrado: true,
          estado: String(data[i][1] || 'Borrador'),
          fechaValidacion: _nomFechaStr(data[i][2]), usuario: String(data[i][3] || ''),
          numOrdenes: parseInt(data[i][4], 10) || 0, totalNeto: parseFloat(data[i][5]) || 0,
          notas: String(data[i][6] || '') };
      }
    }
    return { ok: true, periodo: per, encontrado: false, estado: 'Borrador', numOrdenes: 0, totalNeto: 0 };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
function _nomSetMesEstado(per, o) {
  var sh = _nomMesesSheet();
  var data = sh.getDataRange().getValues();
  var row = [per, o.estado || 'Borrador', o.fechaValidacion || new Date(), o.usuario || '', o.numOrdenes || 0, o.totalNeto || 0, o.notas || ''];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '') === per) { sh.getRange(i + 1, 1, 1, NOM_MESES_HEADERS.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}

// Crea UNA orden de pago (fila CxP en Egresos2026) para un empleado. Idempotente
// por NominaID (columna propia, independiente de RecurrenteID de Gastos Fijos).
// Estructura de fila IDÉNTICA a la que genera Gastos Fijos (_gfAppendCxP): sin
// fecha de pago (col B vacía) + Pagado=false (col N) ⇒ Cuenta por Pagar.
// LockService hace atómico "revisar duplicado + append" (evita doble orden).
function _nomAppendCxP(egSh, iNom1, item, usuario) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'No se pudo obtener el bloqueo, intenta de nuevo' };
  try {
    var monto = parseFloat(String(item.monto || '').replace(/[$,\s]/g, '')) || 0;
    var data = egSh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iNom1 - 1] || '').trim() === item.nominaId) return { ok: false, dup: true, id: item.nominaId };
    }
    var lr = egSh.getLastRow(), lastId = 0;
    if (lr > 1) { var ids = egSh.getRange(2, 1, lr - 1, 1).getValues(); for (var k = 0; k < ids.length; k++) { var n = parseInt(ids[k][0]); if (n > lastId) lastId = n; } }
    var newId = lastId + 1;
    // A ID, B Fecha(''), C Mes, D prioridad, E Proveedor, F Contable, G Tipo, H Subtipo,
    // I Concepto, J Monto, K Notas, L Vencimiento, M/N/O false, P Poliza, Q FormaPago, R Obs, S/T links
    var row = [newId, '', item.periodo, 1, item.proveedor || 'Nómina', item.contable || 'Gasto', 'Fijo', item.subtipo || 'Nómina',
      item.concepto || '', monto, item.notas || '', item.vencimiento || '', false, false, false, '',
      item.formaPago || '', '', '', ''];
    egSh.appendRow(row);
    var newRow = egSh.getLastRow();
    egSh.getRange(newRow, iNom1).setValue(item.nominaId); // NominaID
    var iDiv = _egColEnsure(egSh, 'divisa', 'Divisa');
    egSh.getRange(newRow, iDiv).setValue('MXN');
    try { logAudit(usuario || 'sistema', 'Nómina', 'OrdenCxP', item.nominaId, 'Mes', item.periodo, (item.proveedor || '') + ' ' + monto); } catch (x) {}
  } finally { lock.releaseLock(); }
  return { ok: true, id: newId };
}

// Valida la nómina del mes: crea UNA orden de pago por empleado en Cuentas por
// Pagar (neto CFDI + bonos del mes), y marca el mes como Validada. Idempotente:
// no duplica órdenes (dedup por NominaID); re-enviar con forzar sólo agrega las
// faltantes.
function nominaValidarMes(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) {
      return { ok: false, error: 'Sin autorización para validar nómina (editar_egresos).' };
    }
    var anio = parseInt(body.anio, 10), mes = parseInt(body.mes, 10);
    var per = _nomPeriodo(anio, mes);
    if (!per) return { ok: false, error: 'anio y mes requeridos' };
    var forzar = !!body.forzar;
    var est = nominaMesEstado(anio, mes);
    if (est.ok && est.estado === 'Validada' && !forzar) {
      return { ok: false, yaValidada: true, periodo: per, estado: est,
        error: 'La nómina de ' + per + ' ya fue validada (' + (est.numOrdenes || 0) + ' órdenes). Vuelve a enviar si quieres generar las faltantes.' };
    }
    var data = readNominaMes(anio, mes);
    if (!data.ok) return { ok: false, error: data.error || 'No se pudo leer la nómina del mes.' };
    if (!data.encontrada || !data.empleados.length) {
      return { ok: false, error: 'No hay recibos de nómina (CFDI tipo N) para ' + per + '; no hay nada que validar.' };
    }
    var cfg = _nominaCfg();
    var bonos = _nomBonosDelMes(per);
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egSh = ss.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
    if (!egSh) return { ok: false, error: 'Hoja Egresos2026 no encontrada.' };
    var iNom1 = _egColEnsure(egSh, 'nominaid', 'NominaID');
    var venc = _nomFinDeMes(anio, mes);
    var mesNom = _nomMesNombreEs(mes);
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}

    var creadas = 0, dups = 0, totalNeto = 0, detalle = [];
    for (var i = 0; i < data.empleados.length; i++) {
      var e = data.empleados[i];
      var key1 = String(e.numEmpleado || '').trim(), key2 = String(e.clave || '').trim();
      var bono = 0, seen = {};
      [key1, key2].forEach(function (kk) { if (kk && !seen[kk]) { seen[kk] = 1; bono += (bonos.porNum[kk] || 0); } });
      var netoCfdi = parseFloat(e.neto) || 0;
      var monto = netoCfdi + bono;
      totalNeto += monto;
      var clave = key1 || key2 || ('r' + i);
      var nombre = e.nombre || e.rfc || clave;
      var nominaId = 'NOM-' + per + '-' + String(clave).replace(/[^A-Za-z0-9_\-]/g, '');
      var notas = 'Nómina ' + per + ' · neto CFDI ' + netoCfdi.toFixed(2) + (bono > 0 ? (' + bonos ' + bono.toFixed(2)) : '') + (e.asimilado ? ' · asimilado' : '');
      var r = _nomAppendCxP(egSh, iNom1, {
        periodo: per, nominaId: nominaId,
        proveedor: cfg.cxpProveedor || 'Nómina', contable: 'Gasto', subtipo: 'Nómina',
        concepto: 'Nómina ' + mesNom + ' ' + anio + ' — ' + nombre,
        monto: monto, notas: notas, vencimiento: venc
      }, usuario);
      if (r.ok) { creadas++; detalle.push({ empleado: nombre, monto: monto, id: nominaId }); }
      else if (r.dup) dups++;
    }
    _nomSetMesEstado(per, {
      estado: 'Validada', fechaValidacion: new Date(), usuario: usuario,
      numOrdenes: (est.numOrdenes || 0) + creadas, totalNeto: totalNeto,
      notas: (forzar ? 'Re-generada. ' : '') + creadas + ' orden(es) creadas, ' + dups + ' ya existían.'
    });
    try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (e) {}
    return { ok: true, periodo: per, creadas: creadas, duplicadas: dups, totalNeto: totalNeto,
      numEmpleados: data.empleados.length, bonosTotal: bonos.total, detalle: detalle };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
