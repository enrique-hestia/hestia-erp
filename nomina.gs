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
    cxpCuentaContable: 'Sueldos y salarios',

    // ── Estimado de deducciones (F5) ────────────────────────────────────
    // NO es un motor fiscal. El despacho timbra; el CFDI trae la verdad. Esto
    // sólo sirve para VER el neto aproximado antes de timbrar, y siempre se
    // rotula como "estimado" en la UI. Cuando llega el XML, el real lo pisa.
    estimModo: 'historico',  // 'historico' (ratio real del propio empleado) | 'pct' (fijo)
    estimIsrPct: 10,         // % indicativo sobre percepciones — AJÚSTALO a tu realidad
    estimImssPct: 2.5,       // % indicativo sobre percepciones — AJÚSTALO a tu realidad
    periodicidadDefault: 'quincenal', // si el empleado no la tiene fijada

    // ── Exentos configurables (F5-D) ────────────────────────────────────
    // Los topes en UMA cambian cada año → viven aquí, NO en el código. El ERP no
    // aplica reglas fiscales por su cuenta: mientras topeUMA sea 0 el concepto se
    // muestra como "sin configurar" y no se calcula nada.
    umaPorAnio: {},          // {'2026': 113.14} — lo fija el usuario
    exentos: [
      { clave: 'despensa',    concepto: 'Despensa',          topeUMA: 0, base: 'periodo', notas: '' },
      { clave: 'primaVac',    concepto: 'Prima vacacional',  topeUMA: 0, base: 'anual',   notas: '' },
      { clave: 'aguinaldo',   concepto: 'Aguinaldo',         topeUMA: 0, base: 'anual',   notas: '' },
      { clave: 'gasolina',    concepto: 'Vales de gasolina', topeUMA: 0, base: 'periodo', notas: '' },
      { clave: 'fondoAhorro', concepto: 'Fondo de ahorro',   topeUMA: 0, base: 'anual',   notas: '' },
      { clave: 'prevSocial',  concepto: 'Previsión social',  topeUMA: 0, base: 'anual',   notas: '' }
    ],

    // ── Catálogo de conceptos (F5-2/3) ──────────────────────────────────
    // Nombres tomados del catálogo CFDI de nómina para que EMPATEN con el XML
    // al conciliar. Es CONFIGURABLE: el usuario agrega los que falten sin tocar
    // código. `flag` es sólo una etiqueta (G=gravado, E=exento, PS=previsión
    // social, ''=lo define el despacho); NO dispara ningún cálculo aquí — el
    // gravado/exento exacto llega del CFDI.
    conceptos: [
      { clave: 'sueldo',      concepto: 'Sueldo',                      grupo: 'percepcion', flag: 'G',  activo: true },
      { clave: 'bono',        concepto: 'Bono',                        grupo: 'percepcion', flag: 'G',  activo: true },
      { clave: 'bonoProd',    concepto: 'Bono de Productividad',       grupo: 'percepcion', flag: 'G',  activo: true },
      { clave: 'premioAsis',  concepto: 'Premio de asistencia',        grupo: 'percepcion', flag: 'G',  activo: true },
      { clave: 'premioPunt',  concepto: 'Premio de puntualidad',       grupo: 'percepcion', flag: 'G',  activo: true },
      { clave: 'comisiones',  concepto: 'Comisiones',                  grupo: 'percepcion', flag: 'G',  activo: true },
      { clave: 'despensa',    concepto: 'Despensa',                    grupo: 'percepcion', flag: 'PS', activo: true },
      { clave: 'gasolina',    concepto: 'Vales de gasolina',           grupo: 'percepcion', flag: 'PS', activo: true },
      { clave: 'restaurante', concepto: 'Vales de restaurante',        grupo: 'percepcion', flag: 'PS', activo: true },
      { clave: 'transporte',  concepto: 'Transporte',                  grupo: 'percepcion', flag: 'PS', activo: true },
      { clave: 'fondoAhorro', concepto: 'Fondo de ahorro',             grupo: 'percepcion', flag: 'PS', activo: true },
      { clave: 'prevSocial',  concepto: 'Previsión social',            grupo: 'percepcion', flag: 'PS', activo: true },
      { clave: 'primaVac',    concepto: 'Prima vacacional',            grupo: 'percepcion', flag: '',   activo: true },
      { clave: 'vacPagadas',  concepto: 'Días de vacaciones pagadas',  grupo: 'percepcion', flag: '',   activo: true },
      { clave: 'aguinaldo',   concepto: 'Aguinaldo',                   grupo: 'percepcion', flag: '',   activo: true },
      { clave: 'hrExtraDob',  concepto: 'Horas extras dobles',         grupo: 'percepcion', flag: '',   activo: true },
      { clave: 'hrExtraTri',  concepto: 'Horas extras triples',        grupo: 'percepcion', flag: '',   activo: true },
      { clave: 'primaDom',    concepto: 'Prima dominical',             grupo: 'percepcion', flag: '',   activo: true },
      { clave: 'diasFest',    concepto: 'Días festivos',               grupo: 'percepcion', flag: '',   activo: true },
      { clave: 'isr',         concepto: 'ISR',                         grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'subsEmpleo',  concepto: 'Subsidio para el empleo',     grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'imss',        concepto: 'IMSS',                        grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'infonavit',   concepto: 'INFONAVIT',                   grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'prestamos',   concepto: 'PRESTAMOS',                   grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'fonacot',     concepto: 'FONACOT',                     grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'pensionAlim', concepto: 'PENSION ALIMENTICIA',         grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'anticipos',   concepto: 'ANTICIPO SUELDOS',            grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'cajaAhorro',  concepto: 'CAJA DE AHORRO',              grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'cuotaSind',   concepto: 'CUOTA SINDICAL',              grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'fondoAhorroD',concepto: 'FONDO DE AHORRO',             grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'incapacidad', concepto: 'Descuento por incapacidad',   grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'ausentismo',  concepto: 'Ausencia (Ausentismo)',       grupo: 'deduccion',  flag: '',   activo: true },
      { clave: 'otrosDesc',   concepto: 'OTROS DESCUENTOS',            grupo: 'deduccion',  flag: '',   activo: true }
    ]
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

// ── Libro propio de Nómina ─────────────────────────────────────────────
// Cada módulo en su propio spreadsheet para escalar. Si NOMINA_SS_ID no está
// configurado, cae al libro principal (SHEET_ID) para no romper nada antes de migrar.
function _nomBook() {
  var id = '';
  try { id = PropertiesService.getScriptProperties().getProperty('NOMINA_SS_ID') || ''; } catch (e) {}
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  return SpreadsheetApp.openById(SHEET_ID);
}
// Crea el libro propio de Nómina y COPIA ahí las hojas existentes (las originales
// quedan como respaldo en el libro principal). Correr una vez: setupNominaBook()
function setupNominaBook() {
  try {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('NOMINA_SS_ID'), ss = null;
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
    if (!ss) { ss = SpreadsheetApp.create('Nómina — Hestia Fertility'); props.setProperty('NOMINA_SS_ID', ss.getId()); }
    var main = SpreadsheetApp.openById(SHEET_ID);
    var migradas = [];
    [NOM_EMP_TAB, NOM_MESES_TAB, NOM_BONOS_TAB].forEach(function (tab) {
      var src = main.getSheetByName(tab);
      if (src && !ss.getSheetByName(tab)) { src.copyTo(ss).setName(tab); migradas.push(tab); }
    });
    var def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja1');
    if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch (e) {} }
    return { ok: true, id: ss.getId(), url: ss.getUrl(), migradas: migradas };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── Catálogo de Empleados (hoja "Empleados" en el libro de Nómina) ──────
var NOM_EMP_TAB = 'Empleados';
// 'Departamento' es el "Área" de la hoja del usuario (lo llena también el CFDI).
// 'FechaNacimiento' se agrega en F5 (append-safe: _nomEmpSheet añade lo que falte).
var NOM_EMP_HEADERS = [
  'NumEmpleado', 'Nombre', 'RFC', 'CURP', 'NSS', 'Puesto', 'Departamento',
  'FechaIngreso', 'FechaNacimiento', 'SalarioDiario', 'SBC', 'Periodicidad', 'Tipo',
  'Banco', 'CLABE', 'UsuarioEmail', 'Activo', 'Notas'
];
function _nomEmpSheet() {
  var ss = _nomBook();
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
      if (o.FechaNacimiento instanceof Date) o.FechaNacimiento = Utilities.formatDate(o.FechaNacimiento, Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyy-MM-dd');
      o.Activo = (String(o.Activo).toLowerCase() !== 'no' && String(o.Activo).toLowerCase() !== 'false' && o.Activo !== false);
      o.Periodicidad = _nomPeriodicidadNorm(o.Periodicidad);
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
 *   POST nominaValidarMes {token, anio, mes, forzar}  ⛔ DESACTIVADA — la
 *        validación a CxP es POR PERIODO (nominaValidarPeriodo). Responde
 *        {ok:false} y ya no genera órdenes; ver comentario en la función.
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
  var ss = _nomBook();
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
  var ss = _nomBook();
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

/* ═══════════════════════════════════════════════════════════════════════
 * ⛔ DESACTIVADA — la validación a Cuentas por Pagar es AHORA POR PERIODO.
 *
 * POR QUÉ: convivían dos vías a CxP que NO se deduplicaban entre sí, porque
 * cada una arma un NominaID distinto para el mismo pago:
 *     nominaValidarMes     → 'NOM-2026-07-E01'
 *     nominaValidarPeriodo → 'NOM-QNA-2026-07-1-E01'
 * El dedup de _nomAppendCxP compara NominaID exacto, así que validar el mes
 * DESPUÉS de haber validado sus quincenas creaba órdenes duplicadas y se podía
 * pagar dos veces al mismo empleado. La oficial es POR PERIODO (F5).
 *
 * QUÉ SIGUE VIVO: la vista mensual de CFDI (readNominaMes / nominaMesEstado /
 * nominaBonos) es de LECTURA y no se toca — sirve para ver y conciliar los
 * recibos timbrados del mes. Lo único desactivado es la GENERACIÓN de órdenes.
 *
 * NO SE BORRA: la implementación original queda íntegra abajo, en
 * _nominaValidarMes_LEGACY_DESACTIVADA(), como referencia y por si hubiera que
 * auditar cómo se generaron las órdenes viejas. NO la vuelvas a cablear sin
 * unificar antes el NominaID con el de nominaValidarPeriodo.
 * ═══════════════════════════════════════════════════════════════════════ */
function nominaValidarMes(body) {
  return {
    ok: false,
    desactivada: true,
    error: 'La validación de nómina a Cuentas por Pagar ahora es POR PERIODO, no por mes. '
         + 'Ve a Nómina → «Captura por periodo», elige el periodo (semanal/quincenal/mensual) y usa «✓ Validar y enviar a CxP». '
         + 'La vista mensual de CFDI queda solo para consultar y conciliar los recibos timbrados.'
  };
}

// Implementación original (ver comentario de arriba). Desactivada: ya no se
// llama desde ninguna ruta. Se conserva como referencia histórica.
function _nominaValidarMes_LEGACY_DESACTIVADA(body) {
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

/* ═══════════════════════════════════════════════════════════════════════
 * F5 — Periodos de nómina (semanal/quincenal/mensual, MIXTO por empleado),
 * captura por empleado×periodo, exentos configurables y control de SBC.
 *
 * Qué NO es: un motor fiscal. El despacho del usuario timbra. Aquí el ISR/IMSS
 * que se muestra antes de timbrar es un ESTIMADO rotulado como tal; cuando llega
 * el CFDI (_nomParseCfdi / readNominaMes) el dato REAL lo reemplaza y se calcula
 * la Diferencia. En reportes MANDA el real.
 *
 * Hojas nuevas (libro NOMINA_SS_ID, se crean solas al primer uso):
 *   - "Nomina_Periodos" : PeriodoID | Tipo | FechaInicio | FechaFin | FechaPago |
 *                         Estatus | Notas | CreadoEn | CreadoPor
 *   - "Nomina_Captura"  : una fila por empleado × periodo (percepciones + estimado
 *                         + real del CFDI + diferencia + pago)
 *   - "Nomina_SBC"      : SBC por empleado y bimestre + casilla "Presentado"
 *
 * Rutas (registrar en core.gs GET / finance.gs POST):
 *   GET  nominaPeriodos&anio=&tipo=        → readNominaPeriodos(anio, tipo)
 *   GET  nominaCaptura&periodoId=          → readNominaCaptura(periodoId)
 *   GET  nominaSBC&anio=                   → readNominaSBC(anio)
 *   POST nominaGenerarPeriodos {token, anio, tipo}
 *   POST nominaPeriodoSave     {token, periodo}
 *   POST nominaPeriodoEstatus  {token, periodoId, estatus}
 *   POST saveNominaCaptura     {token, periodoId, filas}
 *   POST nominaConciliarPeriodo{token, periodoId}
 *   POST nominaValidarPeriodo  {token, periodoId, forzar}
 *   POST saveNominaSBC         {token, anio, bimestre, numEmpleado, sbc, presentado, notas}
 * ═══════════════════════════════════════════════════════════════════════ */

var NOM_PERIODICIDADES = ['semanal', 'quincenal', 'mensual'];
var NOM_SBC_BIMESTRES = 6; // IMSS: 6 bimestres al año (ene-feb … nov-dic)

function _nom2(n) { n = parseInt(n, 10) || 0; return (n < 10 ? '0' : '') + n; }
function _nomD2S(d) { return d.getFullYear() + '-' + _nom2(d.getMonth() + 1) + '-' + _nom2(d.getDate()); }
function _nomNum(v) { var x = parseFloat(String(v == null ? 0 : v).replace(/[$,\s]/g, '')); return isNaN(x) ? 0 : x; }
function _nomPeriodicidadNorm(v) {
  var s = String(v || '').trim().toLowerCase();
  if (s.indexOf('quincen') > -1 || s === 'q' || s === '15') return 'quincenal';
  if (s.indexOf('seman') > -1 || s === 's' || s === '7') return 'semanal';
  if (s.indexOf('mens') > -1 || s === 'm' || s === '30') return 'mensual';
  return ''; // vacío = usa el default de config
}
// Días nominales que paga cada periodicidad (para sembrar el sueldo base).
function _nomDiasPeriodicidad(tipo) {
  tipo = _nomPeriodicidadNorm(tipo);
  if (tipo === 'semanal') return 7;
  if (tipo === 'mensual') return 30;
  return 15; // quincenal
}
function _nomCfgSafe() { try { return _nominaCfg(); } catch (e) { return _nominaCfgDefault(); } }

// ── A. Generador de periodos ───────────────────────────────────────────
// Devuelve (sin escribir) todos los periodos de un tipo para un año.
//  - mensual   : MEN-2026-01 … MEN-2026-12  (día 1 → último del mes)
//  - quincenal : QNA-2026-01-1 (1–15) y QNA-2026-01-2 (16–fin) → 24 al año
//  - semanal   : SEM-2026-01 … lunes→domingo, arrancando en el lunes de la
//                semana que contiene el 1-ene (cobertura completa, sin traslape)
function _nomGenPeriodos(tipo, anio) {
  anio = parseInt(anio, 10);
  tipo = _nomPeriodicidadNorm(tipo);
  var out = [];
  if (!anio || !tipo) return out;
  if (tipo === 'mensual') {
    for (var m = 1; m <= 12; m++) {
      var ini = new Date(anio, m - 1, 1), fin = new Date(anio, m, 0);
      out.push({ periodoId: 'MEN-' + anio + '-' + _nom2(m), tipo: 'mensual',
        fechaInicio: _nomD2S(ini), fechaFin: _nomD2S(fin), fechaPago: _nomD2S(fin) });
    }
    return out;
  }
  if (tipo === 'quincenal') {
    for (var q = 1; q <= 12; q++) {
      var f15 = new Date(anio, q - 1, 15), fm = new Date(anio, q, 0);
      out.push({ periodoId: 'QNA-' + anio + '-' + _nom2(q) + '-1', tipo: 'quincenal',
        fechaInicio: _nomD2S(new Date(anio, q - 1, 1)), fechaFin: _nomD2S(f15), fechaPago: _nomD2S(f15) });
      out.push({ periodoId: 'QNA-' + anio + '-' + _nom2(q) + '-2', tipo: 'quincenal',
        fechaInicio: _nomD2S(new Date(anio, q - 1, 16)), fechaFin: _nomD2S(fm), fechaPago: _nomD2S(fm) });
    }
    return out;
  }
  // semanal
  var d0 = new Date(anio, 0, 1), dow = d0.getDay();       // 0=domingo
  var back = (dow === 0 ? 6 : dow - 1);                    // retrocede al lunes
  var cur = new Date(anio, 0, 1 - back), n = 1;
  while (cur.getFullYear() <= anio && n <= 54) {
    var ffin = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 6);
    out.push({ periodoId: 'SEM-' + anio + '-' + _nom2(n), tipo: 'semanal',
      fechaInicio: _nomD2S(cur), fechaFin: _nomD2S(ffin), fechaPago: _nomD2S(ffin) });
    cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 7);
    n++;
  }
  return out;
}

// Crea la hoja con sus headers, o AÑADE (append-safe) los que falten si la hoja
// ya existía de una versión anterior. No reordena ni toca los datos existentes.
function _nomSheetEnsure(tab, headers) {
  var ss = _nomBook();
  var sh = ss.getSheetByName(tab);
  if (!sh) {
    sh = ss.insertSheet(tab);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  var last = Math.max(1, sh.getLastColumn());
  var ex = sh.getRange(1, 1, 1, last).getValues()[0].map(function (v) { return String(v); });
  var add = [];
  for (var i = 0; i < headers.length; i++) if (ex.indexOf(headers[i]) === -1) add.push(headers[i]);
  if (add.length) {
    sh.getRange(1, ex.length + 1, 1, add.length).setValues([add]);
    sh.getRange(1, 1, 1, ex.length + add.length).setFontWeight('bold');
  }
  return sh;
}

var NOM_PER_TAB = 'Nomina_Periodos';
// TipoNomina (ORDINARIA/EXTRAORDINARIA) va al FINAL para no mover las columnas
// posicionales que ya usa el código (Estatus=6, Notas=7).
var NOM_PER_HEADERS = ['PeriodoID', 'Tipo', 'FechaInicio', 'FechaFin', 'FechaPago', 'Estatus', 'Notas', 'CreadoEn', 'CreadoPor', 'TipoNomina'];
var NOM_TIPOS_NOMINA = ['ORDINARIA', 'EXTRAORDINARIA'];
function _nomTipoNominaNorm(v) {
  var s = String(v || '').trim().toUpperCase();
  return (s.indexOf('EXTRA') > -1) ? 'EXTRAORDINARIA' : 'ORDINARIA';
}
function _nomPerSheet() { return _nomSheetEnsure(NOM_PER_TAB, NOM_PER_HEADERS); }
function _nomPerRow2Obj(t) {
  return {
    periodoId: String(t[0] || ''), tipo: _nomPeriodicidadNorm(t[1]) || String(t[1] || ''),
    fechaInicio: _nomFechaStr(t[2]), fechaFin: _nomFechaStr(t[3]), fechaPago: _nomFechaStr(t[4]),
    estatus: String(t[5] || 'borrador').toLowerCase(), notas: String(t[6] || ''),
    creadoEn: _nomFechaStr(t[7]), creadoPor: String(t[8] || ''),
    tipoNomina: _nomTipoNominaNorm(t[9]),
    // Días para cálculo (la plantilla del despacho los llama así): 15 en quincena.
    diasCalculo: _nomDiasPeriodicidad(_nomPeriodicidadNorm(t[1]))
  };
}
function readNominaPeriodos(anio, tipo) {
  try {
    var sh = _nomPerSheet();
    var data = sh.getDataRange().getValues();
    var fa = String(anio || '').trim(), ft = _nomPeriodicidadNorm(tipo);
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!String(data[i][0] || '').trim()) continue;
      var o = _nomPerRow2Obj(data[i]);
      if (fa && String(o.fechaInicio || '').substring(0, 4) !== fa) continue;
      if (ft && o.tipo !== ft) continue;
      out.push(o);
    }
    out.sort(function (a, b) { return String(a.fechaInicio).localeCompare(String(b.fechaInicio)) || String(a.periodoId).localeCompare(String(b.periodoId)); });
    return { ok: true, anio: fa, tipo: ft, periodos: out };
  } catch (ex) { return { ok: false, error: ex.message, periodos: [] }; }
}
function _nomPerGet(periodoId) {
  var sh = _nomPerSheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === String(periodoId || '').trim()) {
      var o = _nomPerRow2Obj(data[i]); o._row = i + 1; return o;
    }
  }
  return null;
}
// Genera (idempotente) los periodos faltantes de un tipo para un año.
function nominaGenerarPeriodos(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var anio = parseInt(body.anio, 10), tipo = _nomPeriodicidadNorm(body.tipo);
    if (!anio) return { ok: false, error: 'anio requerido' };
    if (!tipo) return { ok: false, error: 'tipo debe ser semanal, quincenal o mensual' };
    var gen = _nomGenPeriodos(tipo, anio);
    if (!gen.length) return { ok: false, error: 'No se pudieron generar periodos.' };
    var sh = _nomPerSheet();
    var data = sh.getDataRange().getValues();
    var ex = {};
    for (var i = 1; i < data.length; i++) ex[String(data[i][0] || '').trim()] = 1;
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
    var nuevos = [], ahora = new Date();
    gen.forEach(function (p) {
      if (ex[p.periodoId]) return;
      // El calendario del año es nómina ORDINARIA; las extraordinarias se crean
      // sueltas con "Nuevo periodo".
      nuevos.push([p.periodoId, p.tipo, p.fechaInicio, p.fechaFin, p.fechaPago, 'borrador', '', ahora, usuario, 'ORDINARIA']);
    });
    if (nuevos.length) sh.getRange(sh.getLastRow() + 1, 1, nuevos.length, NOM_PER_HEADERS.length).setValues(nuevos);
    try { logAudit(usuario || 'sistema', 'Nómina', 'GenerarPeriodos', tipo + ' ' + anio, '', '', String(nuevos.length)); } catch (x) {}
    return { ok: true, anio: anio, tipo: tipo, creados: nuevos.length, existentes: gen.length - nuevos.length, total: gen.length };
  } catch (ex2) { return { ok: false, error: ex2.message }; }
}
// Alta/edición manual de UN periodo (botón "Nuevo periodo").
function nominaPeriodoSave(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var p = body.periodo || {};
    var id = String(p.periodoId || '').trim();
    var tipo = _nomPeriodicidadNorm(p.tipo);
    if (!tipo) return { ok: false, error: 'tipo debe ser semanal, quincenal o mensual' };
    var ini = String(p.fechaInicio || '').substring(0, 10), fin = String(p.fechaFin || '').substring(0, 10);
    if (!ini || !fin) return { ok: false, error: 'FechaInicio y FechaFin son obligatorias.' };
    if (fin < ini) return { ok: false, error: 'La fecha final no puede ser anterior a la inicial.' };
    var tipoNom = _nomTipoNominaNorm(p.tipoNomina);
    if (!id) id = (tipoNom === 'EXTRAORDINARIA' ? 'EXT-' : '') + tipo.substring(0, 3).toUpperCase() + '-' + ini.replace(/-/g, '') + '-' + fin.replace(/-/g, '');
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
    var sh = _nomPerSheet();
    var cur = _nomPerGet(id);
    var row = [id, tipo, ini, fin, String(p.fechaPago || fin).substring(0, 10),
      String(p.estatus || (cur ? cur.estatus : 'borrador')).toLowerCase(), String(p.notas || ''),
      cur ? (cur.creadoEn || new Date()) : new Date(), cur ? (cur.creadoPor || usuario) : usuario, tipoNom];
    if (cur) sh.getRange(cur._row, 1, 1, NOM_PER_HEADERS.length).setValues([row]);
    else sh.appendRow(row);
    try { logAudit(usuario || 'sistema', 'Nómina', cur ? 'EditarPeriodo' : 'CrearPeriodo', id, '', '', tipo + ' ' + ini + '→' + fin); } catch (x) {}
    return { ok: true, periodoId: id, creado: !cur };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
// borrador → validada → pagada
function nominaPeriodoEstatus(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var id = String(body.periodoId || '').trim();
    var est = String(body.estatus || '').toLowerCase();
    if (!id) return { ok: false, error: 'periodoId requerido' };
    if (['borrador', 'validada', 'pagada'].indexOf(est) === -1) return { ok: false, error: 'estatus debe ser borrador, validada o pagada' };
    var p = _nomPerGet(id);
    if (!p) return { ok: false, error: 'Periodo no encontrado: ' + id };
    var sh = _nomPerSheet();
    sh.getRange(p._row, 6).setValue(est);
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
    try { logAudit(usuario || 'sistema', 'Nómina', 'EstatusPeriodo', id, p.estatus, est, ''); } catch (x) {}
    return { ok: true, periodoId: id, estatus: est };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── B/C. Captura por empleado × periodo ────────────────────────────────
var NOM_CAP_TAB = 'Nomina_Captura';
// TotalGravado/TotalExento van al final (append-safe). Los llena EXACTO el CFDI
// al conciliar; en captura manual quedan en 0 = "lo define el timbrado".
var NOM_CAP_HEADERS = ['PeriodoID', 'EmpleadoID', 'SueldoBase', 'Bonos', 'ValesDespensa', 'Combustible',
  'PrimaVacacional', 'OtrasPercepciones', 'TotalPercepciones', 'ISR_Estimado', 'IMSS_Estimado',
  'OtrasDeducciones', 'NetoEstimado', 'ISR_Real', 'IMSS_Real', 'NetoReal', 'UUID_CFDI',
  'Diferencia', 'Pagado', 'FechaPago', 'ActualizadoEn', 'TotalGravado', 'TotalExento',
  'TotalRetenciones', 'TotalOtrasDeducciones'];
var NOM_CAP_PERCEP = ['SueldoBase', 'Bonos', 'ValesDespensa', 'Combustible', 'PrimaVacacional', 'OtrasPercepciones'];
function _nomCapSheet() { return _nomSheetEnsure(NOM_CAP_TAB, NOM_CAP_HEADERS); }

// ── Desglose por concepto (gravado/exento) ─────────────────────────────
// Una fila por concepto × empleado × periodo. Hoy lo llena el CFDI al conciliar
// (que trae el gravado/exento EXACTO por concepto, ya parseado en _nomParseCfdi).
var NOM_DET_TAB = 'Nomina_Captura_Det';
var NOM_DET_HEADERS = ['PeriodoID', 'EmpleadoID', 'Grupo', 'Clave', 'Concepto', 'Importe',
  'Gravado', 'Exento', 'Origen', 'ActualizadoEn'];
function _nomDetSheet() { return _nomSheetEnsure(NOM_DET_TAB, NOM_DET_HEADERS); }
// Desglose guardado de un periodo → { empleadoId: [ {..} ] }
function _nomDetDelPeriodo(periodoId) {
  var sh = _nomDetSheet();
  var data = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() !== String(periodoId)) continue;
    var k = String(data[i][1] || '').trim();
    if (!out[k]) out[k] = [];
    out[k].push({ grupo: String(data[i][2] || ''), clave: String(data[i][3] || ''), concepto: String(data[i][4] || ''),
      importe: _nomNum(data[i][5]), gravado: _nomNum(data[i][6]), exento: _nomNum(data[i][7]), origen: String(data[i][8] || '') });
  }
  return out;
}
// Reemplaza el desglose de un empleado/periodo (borra el anterior y escribe el nuevo).
function _nomDetReemplaza(periodoId, empleadoId, lineas, origen) {
  var sh = _nomDetSheet();
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0] || '').trim() === String(periodoId) && String(data[i][1] || '').trim() === String(empleadoId)) sh.deleteRow(i + 1);
  }
  if (!lineas || !lineas.length) return 0;
  var ahora = new Date();
  var rows = lineas.map(function (l) {
    return [periodoId, empleadoId, l.grupo || '', l.clave || '', l.concepto || '',
      _nomNum(l.importe), _nomNum(l.gravado), _nomNum(l.exento), origen || 'CFDI', ahora];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, NOM_DET_HEADERS.length).setValues(rows);
  return rows.length;
}
// Suma de percepciones de una fila.
function _nomCapPercepciones(f) {
  var t = 0;
  for (var i = 0; i < NOM_CAP_PERCEP.length; i++) t += _nomNum(f[NOM_CAP_PERCEP[i]]);
  return t;
}
// Recalcula TotalPercepciones + estimado + neto de UNA fila.
// hist: {isr, imss, n} — ratios reales históricos del propio empleado (opcional).
// El estimado NUNCA pisa al real: si la fila ya tiene UUID_CFDI, NetoReal manda.
function _nomCapCalcFila(f, cfg, hist) {
  cfg = cfg || _nomCfgSafe();
  var perc = _nomCapPercepciones(f);
  f.TotalPercepciones = perc;
  f.OtrasDeducciones = _nomNum(f.OtrasDeducciones);
  var tieneCfdi = !!String(f.UUID_CFDI || '');
  // Una fila YA CONCILIADA CONGELA su estimado: el estimado sólo tiene sentido
  // como "lo que creías ANTES de timbrar", y es contra ése que se mide la
  // Diferencia. Recalcularlo aquí sería circular (el histórico del empleado ya
  // incluye ESTE mismo CFDI) y haría desaparecer la diferencia sola.
  var congelado = tieneCfdi && f._estimGuardado;
  if (congelado) {
    f.ISR_Estimado = _nomNum(f.ISR_Estimado);
    f.IMSS_Estimado = _nomNum(f.IMSS_Estimado);
    f.NetoEstimado = _nomNum(f.NetoEstimado);
    f.estimFuente = 'congelado';
    f.estimIsrPctEfec = perc > 0 ? (f.ISR_Estimado / perc * 100) : 0;
    f.estimImssPctEfec = perc > 0 ? (f.IMSS_Estimado / perc * 100) : 0;
  } else {
    var usaHist = (String(cfg.estimModo || 'historico') === 'historico' && hist && hist.n > 0);
    var isrPct = usaHist ? hist.isr : (_nomNum(cfg.estimIsrPct) / 100);
    var imssPct = usaHist ? hist.imss : (_nomNum(cfg.estimImssPct) / 100);
    f.ISR_Estimado = perc * isrPct;
    f.IMSS_Estimado = perc * imssPct;
    f.NetoEstimado = perc - f.ISR_Estimado - f.IMSS_Estimado - f.OtrasDeducciones;
    f.estimFuente = usaHist ? 'historico' : 'pct';
    // % efectivos usados — el front los reutiliza para recalcular en vivo sin ir al servidor.
    f.estimIsrPctEfec = isrPct * 100;
    f.estimImssPctEfec = imssPct * 100;
  }
  // Diferencia sólo tiene sentido cuando ya hay CFDI.
  f.Diferencia = tieneCfdi ? (_nomNum(f.NetoReal) - f.NetoEstimado) : 0;
  f.conCfdi = tieneCfdi;
  f.NetoFinal = tieneCfdi ? _nomNum(f.NetoReal) : f.NetoEstimado; // el REAL manda
  return f;
}
// Aplica el CFDI a la fila: el real reemplaza al estimado y se calcula Diferencia.
function _nomCapAplicaCfdi(f, r) {
  f.ISR_Real = _nomNum(r.isrRetenido);
  f.IMSS_Real = _nomNum(r.imss);
  f.NetoReal = (r.neto != null) ? _nomNum(r.neto)
    : (_nomNum(r.totalPercepciones) + _nomNum(r.totalOtrosPagos) - _nomNum(r.totalDeducciones));
  f.UUID_CFDI = String(r.uuid || '');
  // Gravado/exento EXACTO del CFDI (el ERP no lo calcula: lo timbró el despacho).
  f.TotalGravado = _nomNum(r.gravado);
  f.TotalExento = _nomNum(r.exento);
  // Estructura de totales de la plantilla: retenciones (ISR+IMSS) vs otras deducciones.
  f.TotalRetenciones = _nomNum(r.isrRetenido) + _nomNum(r.imss);
  f.TotalOtrasDeducciones = _nomNum(r.totalDeducciones) - f.TotalRetenciones;
  if (f.TotalOtrasDeducciones < 0) f.TotalOtrasDeducciones = 0;
  f.Diferencia = _nomNum(f.NetoReal) - _nomNum(f.NetoEstimado);
  f.conCfdi = true;
  f.NetoFinal = _nomNum(f.NetoReal);
  return f;
}
function _nomCapRow2Obj(t, h) {
  var o = {};
  for (var c = 0; c < h.length; c++) o[String(h[c])] = t[c];
  o.PeriodoID = String(o.PeriodoID || '');
  o.EmpleadoID = String(o.EmpleadoID || '');
  o.UUID_CFDI = String(o.UUID_CFDI || '');
  o.Pagado = (o.Pagado === true || String(o.Pagado).toLowerCase() === 'sí' || String(o.Pagado).toLowerCase() === 'si' || String(o.Pagado).toLowerCase() === 'true');
  o.FechaPago = _nomFechaStr(o.FechaPago);
  // Marca que el estimado viene de la hoja (ya se calculó antes de timbrar):
  // _nomCapCalcFila lo CONGELA si la fila ya tiene CFDI, para no recalcularlo
  // con un histórico que ya contiene este mismo recibo.
  o._estimGuardado = (o.NetoEstimado !== '' && o.NetoEstimado !== null && o.NetoEstimado !== undefined);
  return o;
}
// Ratios reales históricos por empleado (sólo filas ya conciliadas con CFDI).
function _nomCapHistorico() {
  var sh = _nomCapSheet();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  var h = data[0].map(function (v) { return String(v); });
  var iEmp = h.indexOf('EmpleadoID'), iPer = h.indexOf('TotalPercepciones'),
      iIsr = h.indexOf('ISR_Real'), iImss = h.indexOf('IMSS_Real'), iUuid = h.indexOf('UUID_CFDI');
  var acc = {};
  for (var i = 1; i < data.length; i++) {
    if (iUuid < 0 || !String(data[i][iUuid] || '')) continue;
    var k = String(data[i][iEmp] || '').trim(); if (!k) continue;
    var p = _nomNum(data[i][iPer]); if (p <= 0) continue;
    if (!acc[k]) acc[k] = { perc: 0, isr: 0, imss: 0, n: 0 };
    acc[k].perc += p; acc[k].isr += _nomNum(data[i][iIsr]); acc[k].imss += _nomNum(data[i][iImss]); acc[k].n++;
  }
  var out = {};
  for (var k2 in acc) if (acc.hasOwnProperty(k2)) {
    var a = acc[k2];
    out[k2] = { isr: a.perc > 0 ? (a.isr / a.perc) : 0, imss: a.perc > 0 ? (a.imss / a.perc) : 0, n: a.n };
  }
  return out;
}
// Empleados que caen en ESTE periodo: los que tienen esa periodicidad (mixto).
function _nomEmpleadosDePeriodicidad(tipo, cfg) {
  var def = _nomPeriodicidadNorm((cfg || _nomCfgSafe()).periodicidadDefault) || 'quincenal';
  var emps = (readEmpleados().empleados || []);
  return emps.filter(function (e) {
    if (!e.Activo) return false;
    var p = _nomPeriodicidadNorm(e.Periodicidad) || def;
    return p === tipo;
  });
}
// Devuelve la captura del periodo: filas guardadas + las que faltan sembradas
// (en memoria, no escribe) para los empleados de esa periodicidad.
function readNominaCaptura(periodoId) {
  try {
    var id = String(periodoId || '').trim();
    if (!id) return { ok: false, error: 'periodoId requerido' };
    var per = _nomPerGet(id);
    if (!per) return { ok: false, error: 'Periodo no encontrado: ' + id };
    var cfg = _nomCfgSafe();
    var sh = _nomCapSheet();
    var data = sh.getDataRange().getValues();
    var h = data[0].map(function (v) { return String(v); });
    var guardadas = {};
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() !== id) continue;
      var o = _nomCapRow2Obj(data[i], h);
      guardadas[o.EmpleadoID] = o;
    }
    var hist = _nomCapHistorico();
    var det = _nomDetDelPeriodo(id);
    var emps = _nomEmpleadosDePeriodicidad(per.tipo, cfg);
    var dias = _nomDiasPeriodicidad(per.tipo);
    var filas = [];
    emps.forEach(function (e) {
      var key = String(e.NumEmpleado || '').trim();
      var f = guardadas[key];
      if (!f) {
        f = { PeriodoID: id, EmpleadoID: key, SueldoBase: _nomNum(e.SalarioDiario) * dias,
          Bonos: 0, ValesDespensa: 0, Combustible: 0, PrimaVacacional: 0, OtrasPercepciones: 0,
          OtrasDeducciones: 0, ISR_Real: 0, IMSS_Real: 0, NetoReal: 0, UUID_CFDI: '',
          Pagado: false, FechaPago: '', _nueva: true };
      }
      // Datos del catálogo (CLAVE=NumEmpleado, CURP/RFC/NSS/PUESTO/DEPTO/SBC).
      f.Nombre = e.Nombre; f.Puesto = e.Puesto; f.Departamento = e.Departamento;
      f.RFC = e.RFC; f.CURP = e.CURP; f.NSS = e.NSS;
      f.Banco = e.Banco; f.CLABE = e.CLABE; f.SBC = _nomNum(e.SBC); f.Tipo = e.Tipo;
      f.SalarioDiario = _nomNum(e.SalarioDiario);
      f.Periodicidad = _nomPeriodicidadNorm(e.Periodicidad) || _nomPeriodicidadNorm(cfg.periodicidadDefault) || 'quincenal';
      _nomCapCalcFila(f, cfg, hist[key]);
      f.detalleCfdi = det[key] || [];
      filas.push(f);
    });
    // Filas guardadas de empleados que ya no están activos / cambiaron de periodicidad:
    // se conservan visibles para no perder histórico.
    for (var k in guardadas) if (guardadas.hasOwnProperty(k)) {
      var yaEsta = false;
      for (var z = 0; z < filas.length; z++) { if (filas[z].EmpleadoID === k) { yaEsta = true; break; } }
      if (yaEsta) continue;
      var hf = guardadas[k]; hf.Nombre = hf.Nombre || k; hf.huerfana = true;
      _nomCapCalcFila(hf, cfg, hist[k]);
      hf.detalleCfdi = det[k] || [];
      filas.push(hf);
    }
    filas.sort(function (a, b) { return String(a.Nombre || '').localeCompare(String(b.Nombre || '')); });
    // Totales con la estructura de la plantilla del despacho: Total de ingresos
    // (gravados/exentos) · Total de retenciones · Total otras deducciones · TOTAL A PAGAR.
    var t = { percepciones: 0, isrEst: 0, imssEst: 0, netoEst: 0, isrReal: 0, imssReal: 0, netoReal: 0,
      netoFinal: 0, diferencia: 0, conCfdi: 0, numEmpleados: filas.length,
      totalIngresos: 0, totalIngresosGravados: 0, totalIngresosExentos: 0,
      totalRetenciones: 0, totalOtrasDeducciones: 0, totalAPagar: 0, totalPagado: 0 };
    filas.forEach(function (f) {
      t.percepciones += f.TotalPercepciones; t.isrEst += f.ISR_Estimado; t.imssEst += f.IMSS_Estimado;
      t.netoEst += f.NetoEstimado; t.netoFinal += f.NetoFinal;
      t.totalIngresos += f.TotalPercepciones;
      t.totalIngresosGravados += _nomNum(f.TotalGravado);
      t.totalIngresosExentos += _nomNum(f.TotalExento);
      t.totalRetenciones += f.conCfdi ? _nomNum(f.TotalRetenciones) : (f.ISR_Estimado + f.IMSS_Estimado);
      t.totalOtrasDeducciones += f.conCfdi ? _nomNum(f.TotalOtrasDeducciones) : _nomNum(f.OtrasDeducciones);
      t.totalAPagar += f.NetoFinal;                      // TOTAL A PAGAR (real si hay CFDI)
      if (f.Pagado) t.totalPagado += f.NetoFinal;        // TOTAL PAGADO
      if (f.conCfdi) { t.conCfdi++; t.isrReal += _nomNum(f.ISR_Real); t.imssReal += _nomNum(f.IMSS_Real); t.netoReal += _nomNum(f.NetoReal); t.diferencia += _nomNum(f.Diferencia); }
    });
    return { ok: true, periodo: per, filas: filas, totales: t, conceptos: cfg.conceptos || [],
      estimModo: cfg.estimModo, estimIsrPct: cfg.estimIsrPct, estimImssPct: cfg.estimImssPct,
      soloLectura: (per.estatus === 'pagada') };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
// Guarda (upsert) las filas capturadas del periodo. Bloquea si ya está pagada.
function saveNominaCaptura(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var id = String(body.periodoId || '').trim();
    if (!id) return { ok: false, error: 'periodoId requerido' };
    var per = _nomPerGet(id);
    if (!per) return { ok: false, error: 'Periodo no encontrado: ' + id };
    if (per.estatus === 'pagada' && !body.forzar) return { ok: false, error: 'El periodo ' + id + ' ya está PAGADO: la captura es de solo lectura.' };
    var filas = body.filas || [];
    if (!filas.length) return { ok: false, error: 'No hay filas que guardar.' };
    var cfg = _nomCfgSafe(), hist = _nomCapHistorico();
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'No se pudo obtener el bloqueo, intenta de nuevo' };
    var guardadas = 0;
    try {
      var sh = _nomCapSheet();
      var data = sh.getDataRange().getValues();
      var h = data[0].map(function (v) { return String(v); });
      var rowByEmp = {};
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() !== id) continue;
        rowByEmp[String(data[i][1] || '').trim()] = i + 1;
      }
      var nuevas = [], actualizadas = 0, ahora = new Date();
      filas.forEach(function (inc) {
        var emp = String(inc.EmpleadoID || '').trim();
        if (!emp) return;
        var prev = rowByEmp[emp] ? _nomCapRow2Obj(data[rowByEmp[emp] - 1], h) : {};
        var f = {
          PeriodoID: id, EmpleadoID: emp,
          SueldoBase: _nomNum(inc.SueldoBase), Bonos: _nomNum(inc.Bonos),
          ValesDespensa: _nomNum(inc.ValesDespensa), Combustible: _nomNum(inc.Combustible),
          PrimaVacacional: _nomNum(inc.PrimaVacacional), OtrasPercepciones: _nomNum(inc.OtrasPercepciones),
          OtrasDeducciones: _nomNum(inc.OtrasDeducciones),
          // El real NO se captura a mano: viene del CFDI (se conserva si ya estaba).
          ISR_Real: _nomNum(prev.ISR_Real), IMSS_Real: _nomNum(prev.IMSS_Real),
          NetoReal: _nomNum(prev.NetoReal), UUID_CFDI: String(prev.UUID_CFDI || ''),
          // Si la fila ya tiene CFDI, se arrastra el estimado ORIGINAL para que
          // _nomCapCalcFila lo congele (la Diferencia se mide contra ése).
          ISR_Estimado: prev.ISR_Estimado, IMSS_Estimado: prev.IMSS_Estimado,
          NetoEstimado: prev.NetoEstimado, _estimGuardado: !!prev._estimGuardado,
          // Gravado/exento y el corte retenciones/otras deducciones los fija el
          // CFDI: no se capturan a mano, se conservan.
          TotalGravado: _nomNum(prev.TotalGravado), TotalExento: _nomNum(prev.TotalExento),
          TotalRetenciones: _nomNum(prev.TotalRetenciones), TotalOtrasDeducciones: _nomNum(prev.TotalOtrasDeducciones),
          Pagado: (inc.Pagado === undefined) ? !!prev.Pagado : !!inc.Pagado,
          FechaPago: String(inc.FechaPago || prev.FechaPago || '').substring(0, 10)
        };
        _nomCapCalcFila(f, cfg, hist[emp]);
        var row = h.map(function (col) {
          if (col === 'ActualizadoEn') return ahora;
          var v = f[col];
          return (v === undefined || v === null) ? '' : v;
        });
        if (rowByEmp[emp]) { sh.getRange(rowByEmp[emp], 1, 1, h.length).setValues([row]); actualizadas++; }
        else nuevas.push(row);
        guardadas++;
      });
      if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, h.length).setValues(nuevas);
      var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
      try { logAudit(usuario || 'sistema', 'Nómina', 'GuardarCaptura', id, '', '', nuevas.length + ' nuevas, ' + actualizadas + ' act.'); } catch (x) {}
    } finally { lock.releaseLock(); }
    return { ok: true, periodoId: id, guardadas: guardadas };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
// Trae los CFDI que cubren el periodo y REEMPLAZA el estimado con el dato real.
// Avisa de las diferencias contra lo estimado.
function nominaConciliarPeriodo(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var id = String(body.periodoId || '').trim();
    if (!id) return { ok: false, error: 'periodoId requerido' };
    var per = _nomPerGet(id);
    if (!per) return { ok: false, error: 'Periodo no encontrado: ' + id };
    // Meses que toca el periodo (una quincena vive en 1 mes; una semana puede cruzar 2).
    var meses = {}, ini = String(per.fechaInicio || ''), fin = String(per.fechaFin || per.fechaInicio || '');
    [ini, fin].forEach(function (d) { if (d) meses[d.substring(0, 7)] = 1; });
    var recibos = [], errores = [];
    for (var ym in meses) if (meses.hasOwnProperty(ym)) {
      var a = parseInt(ym.substring(0, 4), 10), m = parseInt(ym.substring(5, 7), 10);
      var r = readNominaMes(a, m);
      if (!r.ok) { errores.push(ym + ': ' + r.error); continue; }
      (r.recibos || []).forEach(function (x) { recibos.push(x); });
    }
    // Sólo los recibos cuyo periodo de pago cae dentro del periodo capturado.
    // Si el CFDI no trae FechaInicialPago, se usa FechaPago.
    var enRango = recibos.filter(function (r) {
      var d = String(r.fechaInicial || r.fechaPago || r.fecha || '').substring(0, 10);
      if (!d) return false;
      return d >= ini && d <= fin;
    });
    // Agrupa por empleado (varios recibos en el mismo periodo se suman).
    var porEmp = {};
    enRango.forEach(function (r) {
      var k = String(r.numEmpleado || '').trim() || String(r.rfc || '').trim();
      if (!k) return;
      if (!porEmp[k]) porEmp[k] = { isrRetenido: 0, imss: 0, neto: 0, totalPercepciones: 0, totalDeducciones: 0,
        totalOtrosPagos: 0, gravado: 0, exento: 0, uuid: '', n: 0, lineas: [] };
      var e = porEmp[k];
      e.isrRetenido += _nomNum(r.isrRetenido); e.imss += _nomNum(r.imss); e.neto += _nomNum(r.neto);
      e.totalPercepciones += _nomNum(r.totalPercepciones); e.totalDeducciones += _nomNum(r.totalDeducciones);
      e.totalOtrosPagos += _nomNum(r.totalOtrosPagos);
      e.gravado += _nomNum(r.gravado); e.exento += _nomNum(r.exento);
      e.uuid = e.uuid ? (e.uuid + ' ' + r.uuid) : String(r.uuid || ''); e.n++;
      // Desglose por concepto con su gravado/exento tal cual viene del XML.
      (r.percepciones || []).forEach(function (p) {
        e.lineas.push({ grupo: 'percepcion', clave: p.clave || p.tipo, concepto: p.concepto,
          importe: _nomNum(p.importe), gravado: _nomNum(p.gravado), exento: _nomNum(p.exento) });
      });
      (r.deducciones || []).forEach(function (dd) {
        e.lineas.push({ grupo: 'deduccion', clave: dd.clave || dd.tipo, concepto: dd.concepto, importe: _nomNum(dd.importe), gravado: 0, exento: 0 });
      });
      (r.otrosPagos || []).forEach(function (op) {
        e.lineas.push({ grupo: 'otroPago', clave: op.clave || op.tipo, concepto: op.concepto, importe: _nomNum(op.importe), gravado: 0, exento: 0 });
      });
    });
    var cur = readNominaCaptura(id);
    if (!cur.ok) return cur;
    var cfg = _nomCfgSafe(), hist = _nomCapHistorico();
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'No se pudo obtener el bloqueo, intenta de nuevo' };
    var conciliadas = 0, sinCfdi = 0, difs = [], sumDif = 0;
    try {
      var sh = _nomCapSheet();
      var data = sh.getDataRange().getValues();
      var h = data[0].map(function (v) { return String(v); });
      var rowByEmp = {};
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() !== id) continue;
        rowByEmp[String(data[i][1] || '').trim()] = i + 1;
      }
      var nuevas = [], ahora = new Date();
      cur.filas.forEach(function (f) {
        var k = String(f.EmpleadoID || '').trim();
        var r = porEmp[k];
        if (!r) { sinCfdi++; return; }
        _nomCapCalcFila(f, cfg, hist[k]);   // fija el estimado (se congela al conciliar)
        var estAntes = f.NetoEstimado;
        _nomCapAplicaCfdi(f, r);            // el REAL pisa al estimado
        try { _nomDetReemplaza(id, k, r.lineas, 'CFDI'); } catch (xd) {} // desglose gravado/exento
        conciliadas++;
        if (Math.abs(f.Diferencia) >= 0.01) {
          sumDif += f.Diferencia;
          difs.push({ empleado: f.Nombre || k, numEmpleado: k, estimado: estAntes, real: f.NetoReal, diferencia: f.Diferencia });
        }
        var row = h.map(function (col) {
          if (col === 'ActualizadoEn') return ahora;
          var v = f[col];
          return (v === undefined || v === null) ? '' : v;
        });
        if (rowByEmp[k]) sh.getRange(rowByEmp[k], 1, 1, h.length).setValues([row]);
        else nuevas.push(row);
      });
      if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, h.length).setValues(nuevas);
      var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
      try { logAudit(usuario || 'sistema', 'Nómina', 'ConciliarPeriodo', id, '', '', conciliadas + ' con CFDI, ' + difs.length + ' con diferencia'); } catch (x) {}
    } finally { lock.releaseLock(); }
    return { ok: true, periodoId: id, conciliadas: conciliadas, sinCfdi: sinCfdi,
      recibosEnRango: enRango.length, diferencias: difs, sumaDiferencias: sumDif, errores: errores };
  } catch (ex) { return { ok: false, error: ex.message }; }
}
// Valida el periodo capturado → una orden de pago por empleado en CxP.
// Usa el NETO REAL si ya hay CFDI; si no, el estimado (y lo dice en las notas).
// Idempotente por NominaID = NOM-<periodoId>-<empleado>, igual que el flujo mensual.
function nominaValidarPeriodo(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización para validar nómina (editar_egresos).' };
    var id = String(body.periodoId || '').trim();
    if (!id) return { ok: false, error: 'periodoId requerido' };
    var per = _nomPerGet(id);
    if (!per) return { ok: false, error: 'Periodo no encontrado: ' + id };
    if (per.estatus !== 'borrador' && !body.forzar) {
      return { ok: false, yaValidada: true, periodoId: id, estatus: per.estatus,
        error: 'El periodo ' + id + ' ya está ' + per.estatus + '. Vuelve a enviar si quieres generar las órdenes faltantes.' };
    }
    var cap = readNominaCaptura(id);
    if (!cap.ok) return cap;
    if (!cap.filas.length) return { ok: false, error: 'No hay empleados capturados en ' + id + '.' };
    var cfg = _nomCfgSafe();
    var ss = SpreadsheetApp.openById(EGRESOS_SS_2026);
    var egSh = ss.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
    if (!egSh) return { ok: false, error: 'Hoja Egresos2026 no encontrada.' };
    var iNom1 = _egColEnsure(egSh, 'nominaid', 'NominaID');
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
    var creadas = 0, dups = 0, total = 0, detalle = [], estimadas = 0;
    cap.filas.forEach(function (f) {
      var monto = _nomNum(f.NetoFinal);
      if (monto <= 0) return;
      total += monto;
      if (!f.conCfdi) estimadas++;
      var clave = String(f.EmpleadoID || '').replace(/[^A-Za-z0-9_\-]/g, '');
      var nominaId = 'NOM-' + id + '-' + clave;
      var notas = 'Nómina ' + id + ' (' + per.tipo + ' ' + per.fechaInicio + '→' + per.fechaFin + ') · '
        + (f.conCfdi ? ('neto CFDI real ' + monto.toFixed(2)) : ('NETO ESTIMADO ' + monto.toFixed(2) + ' — pendiente de CFDI'));
      var r = _nomAppendCxP(egSh, iNom1, {
        periodo: String(per.fechaPago || per.fechaFin || '').substring(0, 7),
        nominaId: nominaId, proveedor: cfg.cxpProveedor || 'Nómina', contable: 'Gasto', subtipo: 'Nómina',
        concepto: 'Nómina ' + id + ' — ' + (f.Nombre || clave),
        monto: monto, notas: notas, vencimiento: per.fechaPago || per.fechaFin
      }, usuario);
      if (r.ok) { creadas++; detalle.push({ empleado: f.Nombre || clave, monto: monto, id: nominaId, estimado: !f.conCfdi }); }
      else if (r.dup) dups++;
    });
    var shP = _nomPerSheet();
    shP.getRange(per._row, 6).setValue('validada');
    shP.getRange(per._row, 7).setValue((creadas + dups) + ' orden(es) en CxP · ' + creadas + ' nuevas' + (estimadas ? (' · ' + estimadas + ' con neto ESTIMADO') : ''));
    try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (e) {}
    return { ok: true, periodoId: id, creadas: creadas, duplicadas: dups, total: total,
      estimadas: estimadas, numEmpleados: cap.filas.length, detalle: detalle };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// ── E. SBC + control bimestral ─────────────────────────────────────────
var NOM_SBC_TAB = 'Nomina_SBC';
var NOM_SBC_HEADERS = ['Anio', 'Bimestre', 'NumEmpleado', 'SBC', 'Presentado', 'FechaPresentado', 'Usuario', 'Notas'];
function _nomSbcSheet() {
  var ss = _nomBook();
  var sh = ss.getSheetByName(NOM_SBC_TAB);
  if (!sh) {
    sh = ss.insertSheet(NOM_SBC_TAB);
    sh.getRange(1, 1, 1, NOM_SBC_HEADERS.length).setValues([NOM_SBC_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function _nomSbcKey(anio, bim, num) { return String(anio) + '|' + String(bim) + '|' + String(num).trim(); }
// Tablero: empleados × bimestres del año, con SBC y casilla "Presentado".
function readNominaSBC(anio) {
  try {
    var a = parseInt(anio, 10) || (new Date()).getFullYear();
    var sh = _nomSbcSheet();
    var data = sh.getDataRange().getValues();
    var map = {};
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0], 10) !== a) continue;
      var bim = parseInt(data[i][1], 10), num = String(data[i][2] || '').trim();
      if (!bim || !num) continue;
      map[_nomSbcKey(a, bim, num)] = {
        sbc: _nomNum(data[i][3]),
        presentado: (data[i][4] === true || String(data[i][4]).toLowerCase() === 'sí' || String(data[i][4]).toLowerCase() === 'si' || String(data[i][4]).toLowerCase() === 'true'),
        fechaPresentado: _nomFechaStr(data[i][5]), usuario: String(data[i][6] || ''), notas: String(data[i][7] || '')
      };
    }
    var emps = (readEmpleados().empleados || []).filter(function (e) { return e.Activo; });
    var filas = emps.map(function (e) {
      var num = String(e.NumEmpleado || '').trim();
      var bims = [];
      for (var b = 1; b <= NOM_SBC_BIMESTRES; b++) {
        var c = map[_nomSbcKey(a, b, num)];
        bims.push({ bimestre: b, sbc: c ? c.sbc : _nomNum(e.SBC), presentado: c ? c.presentado : false,
          fechaPresentado: c ? c.fechaPresentado : '', notas: c ? c.notas : '', capturado: !!c });
      }
      return { numEmpleado: num, nombre: e.Nombre, puesto: e.Puesto, departamento: e.Departamento,
        sbcCatalogo: _nomNum(e.SBC), periodicidad: e.Periodicidad, bimestres: bims };
    });
    var pend = 0, pres = 0;
    filas.forEach(function (f) { f.bimestres.forEach(function (b) { if (b.presentado) pres++; else pend++; }); });
    return { ok: true, anio: a, numBimestres: NOM_SBC_BIMESTRES, filas: filas,
      totales: { presentados: pres, pendientes: pend, empleados: filas.length } };
  } catch (ex) { return { ok: false, error: ex.message, filas: [] }; }
}
function saveNominaSBC(body) {
  try {
    if (!_tokenHasPermission(body.token || '', 'editar_egresos')) return { ok: false, error: 'Sin autorización (editar_egresos).' };
    var a = parseInt(body.anio, 10), b = parseInt(body.bimestre, 10);
    var num = String(body.numEmpleado || '').trim();
    if (!a || !b || !num) return { ok: false, error: 'anio, bimestre y numEmpleado requeridos' };
    if (b < 1 || b > NOM_SBC_BIMESTRES) return { ok: false, error: 'bimestre debe ser 1–' + NOM_SBC_BIMESTRES };
    var presentado = !!body.presentado;
    var usuario = ''; try { usuario = verifyToken(body.token || '') || ''; } catch (e) {}
    var sh = _nomSbcSheet();
    var data = sh.getDataRange().getValues();
    var found = -1;
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0], 10) === a && parseInt(data[i][1], 10) === b && String(data[i][2] || '').trim() === num) { found = i + 1; break; }
    }
    var prev = found > 0 ? data[found - 1] : null;
    var sbc = (body.sbc === undefined || body.sbc === null || body.sbc === '') ? (prev ? _nomNum(prev[3]) : 0) : _nomNum(body.sbc);
    var row = [a, b, num, sbc, presentado, presentado ? (String(body.fechaPresentado || '').substring(0, 10) || new Date()) : '', usuario, String(body.notas || (prev ? prev[7] : '') || '')];
    if (found > 0) sh.getRange(found, 1, 1, NOM_SBC_HEADERS.length).setValues([row]);
    else sh.appendRow(row);
    try { logAudit(usuario || 'sistema', 'Nómina', 'SBC', num + ' ' + a + '-B' + b, '', '', 'SBC ' + sbc + (presentado ? ' · presentado' : ' · pendiente')); } catch (x) {}
    return { ok: true, anio: a, bimestre: b, numEmpleado: num, sbc: sbc, presentado: presentado };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

// Crea las hojas nuevas de F5 y genera los periodos del año en curso.
// Correr UNA vez desde el editor de Apps Script tras desplegar.
function setupNominaPeriodos() {
  try {
    _nomPerSheet(); _nomCapSheet(); _nomDetSheet(); _nomSbcSheet(); _nomEmpSheet();
    var anio = (new Date()).getFullYear();
    var sh = _nomPerSheet();
    var data = sh.getDataRange().getValues();
    var ex = {};
    for (var i = 1; i < data.length; i++) ex[String(data[i][0] || '').trim()] = 1;
    var nuevos = [], ahora = new Date();
    NOM_PERIODICIDADES.forEach(function (t) {
      _nomGenPeriodos(t, anio).forEach(function (p) {
        if (ex[p.periodoId]) return;
        nuevos.push([p.periodoId, p.tipo, p.fechaInicio, p.fechaFin, p.fechaPago, 'borrador', '', ahora, 'setup', 'ORDINARIA']);
      });
    });
    if (nuevos.length) sh.getRange(sh.getLastRow() + 1, 1, nuevos.length, NOM_PER_HEADERS.length).setValues(nuevos);
    return { ok: true, anio: anio, periodosCreados: nuevos.length,
      hojas: [NOM_PER_TAB, NOM_CAP_TAB, NOM_DET_TAB, NOM_SBC_TAB], libro: _nomBook().getUrl() };
  } catch (ex2) { return { ok: false, error: ex2.message }; }
}
