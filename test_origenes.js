/* Pruebas de la jerarquía de Orígenes + la NO-regresión de summary/presupuesto.
   Carga el CÓDIGO REAL (.gs) en un contexto vm — nada reimplementado. */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const REPO = process.argv[2] || 'G:/Mi unidad/ERP/.claude/worktrees/agent-ac3f5c9747a3fc3ed';
const MUT = process.env.MUT || '';   // mutante activo (para el arnés de mutación)

function src(f){
  let s = fs.readFileSync(path.join(REPO, f), 'utf8').replace(/^﻿/, '');
  if (MUT) s = applyMutant(s, f);
  return s;
}

/* ── Mutantes: rompen UN arreglo cada uno. La prueba correspondiente DEBE fallar. */
function applyMutant(s, f){
  const m = MUT;
  if (f === 'origenes.gs') {
    // M1: el padre puede ser de cualquier tipo → jerarquía de 2 niveles pasa.
    if (m === 'M1') s = s.replace(
      /if \(!_origTipoEsGrupo\(pr\.tipo\)\)\r?\n\s*return 'El padre debe ser de tipo Grupo[^;]*;/,
      '');
    // M2: un Grupo SÍ puede tener padre (ciclo de 2 niveles).
    if (m === 'M2') s = s.replace(
      /if \(tipoCanon === 'Grupo' && padreId\)\r?\n\s*return 'Un Grupo no puede pertenecer[^;]*;/,
      '');
    // M3: ciclo directo (ser su propio padre) permitido.
    if (m === 'M3') s = s.replace(
      /if \(padreId === id\) return 'Un origen no puede ser su propio Grupo\.';/,
      '');
    // M4: grupoId ignora al padre → "el grupo de Lozano" siempre es él mismo.
    if (m === 'M4') s = s.replace(/o\.grupoId = p \|\| id;/, 'o.grupoId = id;');
    // M5: la migración escribe SIN aplicar:true.
    if (m === 'M5') s = s.replace(/var aplicar = \(opts\.aplicar === true\);/, 'var aplicar = true;');
    // M6: _origTipoCanon manda Coordinador/Grupo a "Médico externo" (bug viejo).
    if (m === 'M6') s = s.replace(
      /if \(n\.indexOf\('coordin'\) > -1\) return 'Coordinador';\r?\n\s*if \(n\.indexOf\('grupo'\)\s*> -1\) return 'Grupo';/,
      '');
  }
  if (f === 'summary.gs') {
    // M7: el bucket de Externos se decide por el ORIGEN (el error que se teme).
    if (m === 'M7') s = s.replace(
      /if \(_esExt\) \{\r?\n\s*var _org = String\(r\.origen\|\|''\)\.trim\(\);/,
      "if (_esExt || String(r.origen||'').trim()) {\n          var _org = String(r.origen||'').trim();");
  }
  if (f === 'presupuesto.gs') {
    // M8: presupuesto abre el sub-nivel por origen para CUALQUIER grupo.
    if (m === 'M8') s = s.replace(
      /var s = \(g === 'Agencias'\) \? \(_org \|\| _l2 \|\| g\)\r?\n\s*: \(g === 'Externos'\) \? \(_org \|\| 'Externos — sin atribuir'\)\r?\n\s*: \(_l2 \|\| g\);/,
      "var s = (g === 'Agencias') ? (_org || _l2 || g)\n            : (g === 'Externos') ? (_org || 'Externos — sin atribuir')\n            : (_org || _l2 || g);");
  }
  return s;
}

/* ── Contexto tipo Apps Script ─────────────────────────────────────────── */
function makeCtx(sheets){
  sheets = sheets || {};
  const ctx = {
    console, JSON, Math, Date, String, Number, Object, Array, parseInt, parseFloat, isNaN, RegExp, Error,
    SHEET_ID: 'SS_MAIN', PACIENTES_SS_ID: 'SS_PAC',
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: { getScriptCache: () => ({ remove(){}, get(){return null;}, put(){} }) },
    Session: { getScriptTimeZone: () => 'America/Mexico_City' },
    Utilities: { formatDate: (d) => d.toISOString().substring(0,10) },
    logAudit: () => {},
    _tokenHasPermission: () => true,
    jsonResponse: (o) => o,
    SpreadsheetApp: {
      openById(id){
        const book = sheets[id];
        if (!book) throw new Error('No existe el libro ' + id);
        return {
          getSheetByName: (n) => book[n] || null,
          getSheets: () => Object.keys(book).map(k => book[k]),
          insertSheet: (n) => (book[n] = mkSheet(n, [])),
        };
      }
    }
  };
  ctx.globalThis = ctx;
  return vm.createContext(ctx);
}

/* Hoja de cálculo falsa, con la semántica que importa: getRange/setValue,
   getLastRow/getLastColumn, appendRow. Los datos viven en `rows`. */
function mkSheet(name, rows){
  const sh = {
    name, rows,
    getName: () => name,
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    getLastRow: () => rows.length,
    getLastColumn: () => rows.reduce((m, r) => Math.max(m, r.length), 0),
    setFrozenRows(){ return sh; },
    getRange(r, c, nr, nc){
      nr = nr || 1; nc = nc || 1;
      return {
        getValues(){
          const out = [];
          for (let i = 0; i < nr; i++){
            const row = rows[r - 1 + i] || [];
            const line = [];
            for (let j = 0; j < nc; j++) line.push(row[c - 1 + j] === undefined ? '' : row[c - 1 + j]);
            out.push(line);
          }
          return out;
        },
        setValues(vals){
          for (let i = 0; i < vals.length; i++){
            while (rows.length < r + i) rows.push([]);
            const row = rows[r - 1 + i];
            for (let j = 0; j < vals[i].length; j++) row[c - 1 + j] = vals[i][j];
          }
          return this;
        },
        setValue(v){
          while (rows.length < r) rows.push([]);
          rows[r - 1][c - 1] = v;
          return this;
        },
        setFontWeight(){ return this; }, setBackground(){ return this; },
        setFontColor(){ return this; }, clearContent(){ return this; },
        setDataValidation(){ return this; },
      };
    },
    appendRow(v){ rows.push(v.slice()); return sh; },
    deleteRow(n){ rows.splice(n - 1, 1); return sh; },
  };
  return sh;
}

/* ── Runner ─────────────────────────────────────────────────────────────── */
const results = [];
function t(name, fn){
  try { fn(); results.push({ name, pass: true }); }
  catch (e) { results.push({ name, pass: false, err: e.message }); }
}
function eq(a, b, msg){
  const A = JSON.stringify(a), B = JSON.stringify(b);
  if (A !== B) throw new Error((msg || '') + ' esperado ' + B + ' pero fue ' + A);
}
function ok(c, msg){ if (!c) throw new Error(msg || 'falso'); }

/* ═══ Catálogo de ejemplo (el del contrato) ═══ */
const CAT_HDR = ['ID','Nombre','Tipo','Alias','Activo','Notas','CreadoEn','Padre'];
const CAT_ROWS = () => [
  CAT_HDR.slice(),
  ['ORIG-00001','Grupo Médico','Grupo','','TRUE','','2026-01-01',''],
  ['ORIG-00002','Dr. Paladino','Médico externo','paladino','TRUE','','2026-01-01','ORIG-00001'],
  ['ORIG-00003','Dr. Lozano','Médico externo','','TRUE','','2026-01-01',''],
  ['ORIG-00004','REPROVIDA','Agencia','repro vida','TRUE','','2026-01-01',''],
  ['ORIG-00005','Daniel Madero','Coordinador','','TRUE','','2026-01-01','ORIG-00001'],
];

function loadOrigenes(catRows, pacRows){
  const books = {
    SS_MAIN: { Origenes_Externos: mkSheet('Origenes_Externos', catRows || CAT_ROWS()) },
    SS_PAC:  { Pacientes: mkSheet('Pacientes', pacRows || [['ID','Nombre Completo del Paciente']]) },
  };
  const ctx = makeCtx(books);
  vm.runInContext(src('origenes.gs'), ctx);
  return { ctx, books };
}

/* ═══════════ 1. _origResolver ═══════════ */
t('_origResolver: médico dentro de un grupo → grupoId = el grupo', () => {
  const { ctx } = loadOrigenes();
  const r = ctx._origResolver('Dr. Paladino');
  eq(r, { id:'ORIG-00002', nombre:'Dr. Paladino', tipo:'Médico externo', padreId:'ORIG-00001', grupoId:'ORIG-00001' });
});
t('_origResolver: médico independiente → grupoId = él mismo', () => {
  const { ctx } = loadOrigenes();
  const r = ctx._origResolver('Dr. Lozano');
  eq(r, { id:'ORIG-00003', nombre:'Dr. Lozano', tipo:'Médico externo', padreId:'', grupoId:'ORIG-00003' });
});
t('_origResolver: agencia → grupoId = ella misma', () => {
  const { ctx } = loadOrigenes();
  eq(ctx._origResolver('REPROVIDA').grupoId, 'ORIG-00004');
  eq(ctx._origResolver('REPROVIDA').tipo, 'Agencia');
});
t('_origResolver: coordinador dentro de un grupo', () => {
  const { ctx } = loadOrigenes();
  const r = ctx._origResolver('Daniel Madero');
  eq(r.tipo, 'Coordinador');
  eq(r.grupoId, 'ORIG-00001');
});
t('_origResolver: nombre inexistente → objeto vacío (no revienta)', () => {
  const { ctx } = loadOrigenes();
  eq(ctx._origResolver('Dr. Quien Sabe'), { id:'', nombre:'', tipo:'', padreId:'', grupoId:'' });
  eq(ctx._origResolver(''),     { id:'', nombre:'', tipo:'', padreId:'', grupoId:'' });
  eq(ctx._origResolver(null),   { id:'', nombre:'', tipo:'', padreId:'', grupoId:'' });
});
t('_origResolver: resuelve por ID y por alias, e ignora acentos/mayúsculas', () => {
  const { ctx } = loadOrigenes();
  eq(ctx._origResolver('ORIG-00002').nombre, 'Dr. Paladino');
  eq(ctx._origResolver('paladino').nombre,   'Dr. Paladino');
  eq(ctx._origResolver('GRUPO MEDICO').id,   'ORIG-00001');   // sin acento, mayúsculas
});

/* ═══════════ 2. Jerarquía: ciclos y 2 niveles ═══════════ */
t('jerarquía: rechaza ciclo directo (ser su propio grupo)', () => {
  const { ctx } = loadOrigenes();
  const cat = ctx._origBuild(CAT_ROWS());
  const err = ctx._origValidaJerarquia(cat, 'ORIG-00002', 'Médico externo', 'ORIG-00002');
  ok(/su propio Grupo/.test(err), 'no rechazó el ciclo directo: ' + err);
});
t('jerarquía: rechaza 2 niveles (padre que no es Grupo)', () => {
  const { ctx } = loadOrigenes();
  const cat = ctx._origBuild(CAT_ROWS());
  // Colgar un médico de OTRO médico = jerarquía de 2 niveles.
  const err = ctx._origValidaJerarquia(cat, 'ORIG-00003', 'Médico externo', 'ORIG-00002');
  ok(/tipo Grupo/.test(err), 'no rechazó el padre no-Grupo: ' + err);
});
t('jerarquía: un Grupo no puede tener padre', () => {
  const { ctx } = loadOrigenes();
  const cat = ctx._origBuild(CAT_ROWS());
  const err = ctx._origValidaJerarquia(cat, '', 'Grupo', 'ORIG-00001');
  ok(/un solo nivel|no puede pertenecer/.test(err), 'no rechazó Grupo-con-padre: ' + err);
});
t('jerarquía: padre inexistente se rechaza', () => {
  const { ctx } = loadOrigenes();
  const cat = ctx._origBuild(CAT_ROWS());
  ok(/no existe/.test(ctx._origValidaJerarquia(cat, '', 'Médico externo', 'ORIG-99999')));
});
t('jerarquía: caso válido pasa (médico → grupo)', () => {
  const { ctx } = loadOrigenes();
  const cat = ctx._origBuild(CAT_ROWS());
  eq(ctx._origValidaJerarquia(cat, 'ORIG-00003', 'Médico externo', 'ORIG-00001'), '');
});
t('jerarquía: datos sucios en la hoja (padre de 2 niveles) se sueltan al leer', () => {
  // Fila metida a mano: un médico colgando de OTRO médico.
  const rows = CAT_ROWS();
  rows.push(['ORIG-00006','Dr. Sucio','Médico externo','','TRUE','','2026-01-01','ORIG-00002']);
  const { ctx } = loadOrigenes(rows);
  const r = ctx._origResolver('Dr. Sucio');
  eq(r.padreId, '', 'el padre no-Grupo debió soltarse');
  eq(r.grupoId, 'ORIG-00006', 'sin grupo válido, su grupo es él mismo');
});
t('saveOrigen: rechaza cambiar de tipo a un Grupo que tiene miembros', () => {
  const { ctx } = loadOrigenes();
  const res = ctx.saveOrigen({ token:'t', id:'ORIG-00001', nombre:'Grupo Médico', tipo:'Agencia' });
  eq(res.ok, false);
  ok(/tienen a este como Grupo/.test(res.error), res.error);
});
t('saveOrigen: alta válida escribe Padre y persiste', () => {
  const { ctx, books } = loadOrigenes();
  const res = ctx.saveOrigen({ token:'t', nombre:'Dra. Nueva', tipo:'Médico externo', padre:'ORIG-00001' });
  eq(res.ok, true);
  const r = ctx._origResolver('Dra. Nueva');
  eq(r.grupoId, 'ORIG-00001');
  eq(r.tipo, 'Médico externo');
});
t('saveOrigen: rechaza el ciclo end-to-end', () => {
  const { ctx } = loadOrigenes();
  const res = ctx.saveOrigen({ token:'t', id:'ORIG-00002', nombre:'Dr. Paladino', tipo:'Médico externo', padre:'ORIG-00002' });
  eq(res.ok, false);
});

/* ═══════════ 3. Compatibilidad con datos VIEJOS (hoja sin Padre) ═══════════ */
t('compat: hoja vieja de 7 columnas (sin Padre) se lee sin romperse', () => {
  const viejo = [
    ['ID','Nombre','Tipo','Alias','Activo','Notas','CreadoEn'],
    ['ORIG-00001','REPROVIDA','Agencia','','TRUE','','2026-01-01'],
    ['ORIG-00002','Dr. Viejo','Médico externo','','TRUE','','2026-01-01'],
  ];
  const { ctx } = loadOrigenes(viejo);
  const reg = ctx.readOrigenes();
  eq(reg.ok, true);
  eq(reg.origenes.length, 2);
  eq(reg.origenes[0].tipo, 'Agencia',        'los "Agencia" viejos deben seguir clasificando igual');
  eq(reg.origenes[1].tipo, 'Médico externo', 'los "Médico externo" viejos deben seguir clasificando igual');
  eq(reg.origenes[0].padre, '');
  eq(reg.origenes[0].grupoId, 'ORIG-00001');
});
t('compat: _origColEnsure agrega Padre AL FINAL, sin desplazar columnas', () => {
  const viejo = [
    ['ID','Nombre','Tipo','Alias','Activo','Notas','CreadoEn'],
    ['ORIG-00001','REPROVIDA','Agencia','','TRUE','','2026-01-01'],
  ];
  const { ctx, books } = loadOrigenes(viejo);
  ctx.saveOrigen({ token:'t', nombre:'Dr. X', tipo:'Médico externo' });
  const hdr = books.SS_MAIN.Origenes_Externos.rows[0];
  eq(hdr[7], 'Padre', 'Padre debe quedar en la col 8 (append)');
  eq(hdr.slice(0,7), ['ID','Nombre','Tipo','Alias','Activo','Notas','CreadoEn'], 'no se desplazó nada');
  eq(books.SS_MAIN.Origenes_Externos.rows[1][1], 'REPROVIDA', 'la fila vieja no se movió');
});
t('_origTipoCanon: los 4 tipos + celda vacía', () => {
  const { ctx } = loadOrigenes();
  eq(ctx._origTipoCanon('Agencia'), 'Agencia');
  eq(ctx._origTipoCanon('Médico externo'), 'Médico externo');
  eq(ctx._origTipoCanon('Grupo'), 'Grupo');
  eq(ctx._origTipoCanon('Coordinador'), 'Coordinador');
  eq(ctx._origTipoCanon(''), 'Médico externo');   // como los clasificaba la vista vieja
});

/* ═══════════ 4. Migración de las fichas ═══════════ */
const PAC_HDR = ['ID','Nombre Completo del Paciente','Fecha de Nacimiento','E-mail','Origen','Canal','Médico Tratante','País','Idioma','Lista'];
const PAC_ROWS = () => [
  PAC_HDR.slice(),
  ['P1','Ana',  '', '', 'Web', 'Grupo Medico', 'Dr. Paladino', 'MX', 'es', 'General'],  // canal sin acento
  ['P2','Bea',  '', '', 'Web', 'REPROVIDA',    '',             'MX', 'es', 'General'],
  ['P3','Cris', '', '', 'Web', 'Clinica Rara', 'Dr. Fantasma', 'MX', 'es', 'General'],  // ninguno empata
  ['P4','Dani', '', '', 'Web', 'Grupo Médico', 'Dr. Lozano',   'MX', 'es', 'General'],  // ya canónico
];
t('migración: SIN aplicar:true no escribe NADA (solo reporte)', () => {
  const { ctx, books } = loadOrigenes(null, PAC_ROWS());
  const antes = JSON.stringify(books.SS_PAC.Pacientes.rows);
  const rep = ctx.migrarCanalesAOrigenes();
  eq(rep.ok, true);
  eq(rep.escritas, 0, 'no debió escribir');
  eq(JSON.stringify(books.SS_PAC.Pacientes.rows), antes, 'la hoja NO debe cambiar');
  ok(rep.nota && /no se escribio nada/.test(rep.nota));
});
t('migración: el reporte separa lo que empata de lo que no', () => {
  const { ctx } = loadOrigenes(null, PAC_ROWS());
  const rep = ctx.migrarCanalesAOrigenes();
  const canal = rep.campos.find(c => c.campo === 'Canal');
  const med   = rep.campos.find(c => c.campo === 'Médico Tratante');
  eq(canal.empatan, 2, 'Grupo Medico + REPROVIDA (Grupo Médico normaliza)');  // "Grupo Medico" y "Grupo Médico" son la MISMA clave normalizada
  eq(canal.noEmpatan, 1, 'Clinica Rara no empata');
  eq(med.empatan, 2, 'Paladino + Lozano');
  eq(med.noEmpatan, 1, 'Dr. Fantasma no empata');
  const fantasma = med.valores.find(v => v.valor === 'Dr. Fantasma');
  eq(fantasma.estado, 'sin_match');
  ok(fantasma.accion, 'debe decirle al usuario qué hacer');
});
t('migración: aplicar:true normaliza al nombre canónico y es idempotente', () => {
  const { ctx, books } = loadOrigenes(null, PAC_ROWS());
  const r1 = ctx.migrarCanalesAOrigenes({ aplicar:true, token:'t' });
  eq(r1.ok, true);
  ok(r1.escritas > 0, 'debió normalizar "Grupo Medico" → "Grupo Médico"');
  const rows = books.SS_PAC.Pacientes.rows;
  eq(rows[1][5], 'Grupo Médico', 'la ficha de Ana quedó canónica');
  // IDEMPOTENTE: la segunda corrida no escribe nada.
  const snap = JSON.stringify(rows);
  const r2 = ctx.migrarCanalesAOrigenes({ aplicar:true, token:'t' });
  eq(r2.escritas, 0, 'la segunda corrida NO debe escribir');
  eq(JSON.stringify(rows), snap, 'la hoja no debe cambiar en la 2a corrida');
});
t('migración: los valores que NO empatan se conservan intactos', () => {
  const { ctx, books } = loadOrigenes(null, PAC_ROWS());
  ctx.migrarCanalesAOrigenes({ aplicar:true, token:'t' });
  const rows = books.SS_PAC.Pacientes.rows;
  eq(rows[3][5], 'Clinica Rara', 'el canal sin match NO se borra');
  eq(rows[3][6], 'Dr. Fantasma', 'el médico sin match NO se borra');
});
t('migración: aplicar:true sin permiso se rechaza', () => {
  const { ctx } = loadOrigenes(null, PAC_ROWS());
  ctx._tokenHasPermission = () => false;
  const r = ctx.migrarCanalesAOrigenes({ aplicar:true, token:'x' });
  eq(r.ok, false);
  ok(/autorizacion/i.test(r.error));
});

/* ═══════════ 5. Opciones de la ficha ═══════════ */
t('ficha: Canal = Grupos+Agencias, Médico = Médicos externos', () => {
  const { ctx } = loadOrigenes();
  const of = ctx._origOpcionesFicha();
  eq(of.ok, true);
  eq(of.canales.sort(), ['Grupo Médico','REPROVIDA']);
  eq(of.medicos.sort(), ['Dr. Lozano','Dr. Paladino']);
  // Daniel Madero (Coordinador) no va en ninguno de los dos selectores...
  ok(of.canales.indexOf('Daniel Madero') < 0 && of.medicos.indexOf('Daniel Madero') < 0);
  // ...pero SÍ viaja en los registros (existe para atribuir ventas).
  ok(of.registros.some(r => r.nombre === 'Daniel Madero' && r.grupoId === 'ORIG-00001'));
});
t('ficha: catálogo vacío → ok:false (core.gs conserva la hoja Opciones)', () => {
  const { ctx } = loadOrigenes([CAT_HDR.slice()]);
  eq(ctx._origOpcionesFicha().ok, false);
});

/* ═══════════════════════════════════════════════════════════════════════
   6. LA PRUEBA QUE MÁS IMPORTA
   Una venta NORMAL atribuida (Diana paga directo, referida por Paladino)
   NO se cuela al bucket de Externos de summary.gs ni de presupuesto.gs.
   Se corre el CÓDIGO REAL con los mismos ingresos, cambiando SOLO el
   campo `origen` de las ventas normales, y se exige salida IDÉNTICA.
   ═══════════════════════════════════════════════════════════════════════ */

/* Filas tal como las devuelve _summaryReadIngresos (contrato interno de summary.gs).
   OJO con el fixture: una línea de Revenue con UN SOLO subgrupo colapsa sus
   subitems a productos y el nivel l2 (el que abre por origen) NO SALE en la
   respuesta. Con un solo producto la prueba sería CIEGA al bug que busca.
   Por eso "FIV" trae DOS categorías y DOS orígenes distintos: así l2 se
   materializa y un cambio de bucket es visible. */
function ingresos(origenEnNormales){
  const O  = origenEnNormales ? 'Dr. Paladino' : '';
  const O2 = origenEnNormales ? 'Dr. Lozano'   : '';
  return [
    // ── Ventas NORMALES (ciclo "FIV"). Diana/Inés pagan directo, referidas.
    { op:'OP-1', fecha:'2026-05-10', fechaRaw:'2026-05-10', paciente:'Diana', categoria:'Tratamientos',
      producto:'FIV', cantidad:1, total:100000, formaPago:'Tarjeta', grupoU:'FIV', origen:O, _anio:2026, _fila:2, _fechaAlt:'' },
    { op:'OP-1b', fecha:'2026-05-14', fechaRaw:'2026-05-14', paciente:'Ines', categoria:'Laboratorio',
      producto:'Congelación', cantidad:1, total:20000, formaPago:'Tarjeta', grupoU:'FIV', origen:O2, _anio:2026, _fila:5, _fechaAlt:'' },
    { op:'OP-2', fecha:'2025-08-10', fechaRaw:'2025-08-10', paciente:'Elsa', categoria:'Tratamientos',
      producto:'FIV', cantidad:1, total:90000, formaPago:'Tarjeta', grupoU:'FIV', origen:O, _anio:2025, _fila:2, _fechaAlt:'' },
    { op:'OP-2b', fecha:'2025-08-14', fechaRaw:'2025-08-14', paciente:'Julia', categoria:'Laboratorio',
      producto:'Congelación', cantidad:1, total:15000, formaPago:'Tarjeta', grupoU:'FIV', origen:O2, _anio:2025, _fila:6, _fechaAlt:'' },
    // ── Venta de EXTERNOS de verdad (ciclo EXTERNOS) — siempre con su origen.
    { op:'OP-3', fecha:'2026-05-11', fechaRaw:'2026-05-11', paciente:'Fabi', categoria:'Externos',
      producto:'Estudio', cantidad:1, total:5000, formaPago:'Efectivo', grupoU:'EXTERNOS', origen:'Dr. Lozano', _anio:2026, _fila:3, _fechaAlt:'' },
    { op:'OP-4', fecha:'2025-08-11', fechaRaw:'2025-08-11', paciente:'Gaby', categoria:'Externos',
      producto:'Estudio', cantidad:1, total:4000, formaPago:'Efectivo', grupoU:'EXTERNOS', origen:'Dr. Lozano', _anio:2025, _fila:3, _fechaAlt:'' },
    // ── Venta de AGENCIA (REPROVIDA).
    { op:'OP-5', fecha:'2026-05-12', fechaRaw:'2026-05-12', paciente:'Hilda', categoria:'REPROVIDA',
      producto:'Ovodonación', cantidad:1, total:200000, formaPago:'Transferencia', grupoU:'REPROVIDA', origen:'REPROVIDA', _anio:2026, _fila:4, _fechaAlt:'' },
  ];
}

function ctxSummary(origenEnNormales){
  const ctx = makeCtx({ SS_MAIN: {} });
  vm.runInContext(src('summary.gs'), ctx);
  vm.runInContext(src('presupuesto.gs'), ctx);
  // Se stubean SOLO los lectores de hoja. La lógica de bucket es la REAL.
  const rows = ingresos(origenEnNormales);
  ctx._summaryReadIngresos = (anio) => rows.filter(r => r._anio === anio);
  ctx._summaryReadEgresos  = () => [];
  ctx.readSummaryConfig    = () => ({ ok:true, rows:[], grupos: ctx.SUMMARY_GRUPOS });
  if (typeof ctx._sumIngresosIds === 'function') ctx._sumIngresosIds = () => ({});
  return ctx;
}

/* Extrae solo el Revenue (líneas + subitems) — que es donde vive el bucket. */
function revenueDe(res){
  ok(res && res.ok, 'readSummary falló: ' + (res && res.error));
  return (res.lineas || [])
    .filter(l => l.grupo === 'REVENUE' && l.tipo === 'dato')
    .map(l => ({
      linea: l.linea,
      actual: Math.round(l.actual * 100) / 100,
      prev: Math.round(l.prev * 100) / 100,
      subs: (l.subitems || []).map(s => ({ label: s.label, actual: Math.round(s.actual * 100) / 100 }))
                              .sort((a, b) => a.label < b.label ? -1 : 1),
    }))
    .sort((a, b) => a.linea < b.linea ? -1 : 1);
}

t('★ summary: atribuir una venta NORMAL no cambia NINGÚN número ni bucket', () => {
  const sin  = revenueDe(ctxSummary(false).readSummary('2026-05-01', '2026-05-31'));
  const con  = revenueDe(ctxSummary(true ).readSummary('2026-05-01', '2026-05-31'));
  // Guarda contra un arnés roto: si no hay Revenue, la prueba no prueba nada.
  ok(sin.length >= 2, 'el arnés no produjo líneas de Revenue: ' + JSON.stringify(sin));
  eq(con, sin, 'el Estado de Resultados cambió al atribuir ventas normales:');
});
t('★ summary: la venta normal atribuida NO aparece bajo Externos/Agencias', () => {
  const con = revenueDe(ctxSummary(true).readSummary('2026-05-01', '2026-05-31'));
  const ext = con.filter(l => /extern|agencia/i.test(l.linea));
  ok(ext.length > 0, 'el arnés debe producir el bucket de Externos/Agencias');
  const total = ext.reduce((s, l) => s + l.actual, 0);
  eq(total, 205000, 'el bucket Externos+Agencias debe traer SOLO OP-3 (5000) y OP-5 (200000)');
  // Y en ningún sub-nivel de externos debe asomarse el FIV de Diana.
  ext.forEach(l => l.subs.forEach(s => {
    ok(!/Paladino/.test(s.label), 'Paladino (venta normal) se coló al bucket de Externos: ' + l.linea + ' › ' + s.label);
  }));
});
t('★ summary: la línea normal (FIV) conserva su sub-nivel por CATEGORÍA, no por origen', () => {
  const con = revenueDe(ctxSummary(true).readSummary('2026-05-01', '2026-05-31'));
  const fiv = con.find(l => l.linea === 'FIV');
  ok(fiv, 'debe existir la línea FIV');
  eq(fiv.actual, 120000);
  // Guarda anti-arnés-ciego: si l2 colapsara, esta prueba no vería nada.
  eq(fiv.subs.map(s => s.label), ['Laboratorio', 'Tratamientos'], 'FIV debe abrirse por CATEGORÍA');
  ok(fiv.subs.every(s => !/Paladino|Lozano/.test(s.label)), 'FIV se abrió por origen: ' + JSON.stringify(fiv.subs));
});

/* ── Presupuesto ── */
function presDe(ctx){
  const r = ctx._presIngresosProy(2026, 3);
  const grupos = (r && r.grupos) ? r.grupos : r;
  return JSON.parse(JSON.stringify(grupos, (k, v) => typeof v === 'number' ? Math.round(v * 100) / 100 : v));
}
t('★ presupuesto: atribuir una venta NORMAL no cambia la proyección', () => {
  const sin = presDe(ctxSummary(false));
  const con = presDe(ctxSummary(true));
  ok(JSON.stringify(sin).length > 50, 'el arnés no produjo proyección: ' + JSON.stringify(sin));
  eq(con, sin, 'la proyección cambió al atribuir ventas normales:');
});
t('★ presupuesto: el grupo FIV no se abre por origen', () => {
  const con = presDe(ctxSummary(true));
  const arr = Array.isArray(con) ? con : (con.grupos || []);
  const fiv = arr.find(g => /fiv/i.test(g.grupo || g.nombre || ''));
  if (fiv) {
    (fiv.subgrupos || []).forEach(s => {
      ok(!/Paladino/.test(s.sub || ''), 'el sub de FIV se abrió por origen: ' + s.sub);
    });
  }
});

/* ── Reporte ── */
let fail = 0;
console.log('\n' + (MUT ? 'MUTANTE ' + MUT : 'CÓDIGO REAL') + '\n' + '─'.repeat(78));
results.forEach(r => {
  console.log((r.pass ? '  PASS  ' : '  FAIL  ') + r.name + (r.pass ? '' : '\n          → ' + r.err));
  if (!r.pass) fail++;
});
console.log('─'.repeat(78));
console.log(`${results.length - fail}/${results.length} pasaron` + (fail ? `, ${fail} fallaron` : ''));
if (results.length === 0) { console.log('ARNÉS ROTO: 0 pruebas'); process.exit(3); }
process.exit(fail ? 1 : 0);
