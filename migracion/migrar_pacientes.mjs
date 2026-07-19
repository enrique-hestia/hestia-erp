// ============================================================================
// VestaOS · Migración ESPEJO de Pacientes (Sheets → Supabase) + RED DE SEGURIDAD
// ----------------------------------------------------------------------------
// Carga los pacientes en la base nueva SIN apagar nada del sistema actual, y
// COMPARA Sheets vs Postgres para probar que cuadran (conteo + campo por campo).
//
// Requisitos: Node 18+ (usa fetch nativo, sin dependencias). Los secretos van
// por VARIABLES DE ENTORNO, nunca en el código ni en el repo:
//   SUPABASE_URL         = https://xxxx.supabase.co
//   SUPABASE_SERVICE_KEY = la llave "service_role" (Project Settings → API).
//                          Ignora la RLS (por eso carga/lee todo). ES SECRETA.
//   TENANT_SLUG          = 'hestia' (opcional; default 'hestia')
//
// PRE: los esquemas core+clinico deben estar EXPUESTOS (Settings → API →
//      Exposed schemas) o PostgREST no ve las tablas.
//
// Uso (exporta antes la hoja Pacientes a CSV: Archivo → Descargar → CSV):
//   node migrar_pacientes.mjs verificar  pacientes.csv   # solo compara (no escribe)
//   node migrar_pacientes.mjs cargar     pacientes.csv   # upsert idempotente
// Flujo recomendado: cargar → verificar (debe dar 0 diferencias).
// ============================================================================

import fs from 'node:fs';

const URL = process.env.SUPABASE_URL;
const KEY = process.env.SUPABASE_SERVICE_KEY;
const SLUG = process.env.TENANT_SLUG || 'hestia';
const [cmd, csvPath] = process.argv.slice(2);

if (!URL || !KEY) { console.error('Falta SUPABASE_URL o SUPABASE_SERVICE_KEY (variables de entorno).'); process.exit(1); }
if (!cmd || !csvPath || !['verificar','cargar'].includes(cmd)) {
  console.error('Uso: node migrar_pacientes.mjs <verificar|cargar> <pacientes.csv>'); process.exit(1);
}
if (!fs.existsSync(csvPath)) { console.error('No existe el CSV: ' + csvPath); process.exit(1); }

// ── PostgREST helper (nunca imprime la llave) ──────────────────────────────
async function rest(method, path, { schema = 'clinico', body, prefer } = {}) {
  const h = {
    apikey: KEY, Authorization: 'Bearer ' + KEY,
    'Accept-Profile': schema, 'Content-Profile': schema,
    'Content-Type': 'application/json'
  };
  if (prefer) h.Prefer = prefer;
  const r = await fetch(URL.replace(/\/$/, '') + '/rest/v1/' + path, { method, headers: h, body: body ? JSON.stringify(body) : undefined });
  const txt = await r.text();
  if (!r.ok) throw new Error('REST ' + r.status + ' ' + path + ' → ' + txt.slice(0, 300));
  return txt ? JSON.parse(txt) : null;
}

// Trae TODAS las filas paginando (PostgREST corta en ~1000 por defecto; sin esto
// la comparación mentiría con "faltan" falsos apenas haya >1000 pacientes).
async function restAllRows(pathBase, schema) {
  const PAGE = 1000; let off = 0, all = [];
  for (;;) {
    const chunk = await rest('GET', pathBase + `&order=id.asc&limit=${PAGE}&offset=${off}`, { schema });
    if (!chunk || !chunk.length) break;
    all = all.concat(chunk);
    if (chunk.length < PAGE) break;
    off += PAGE;
  }
  return all;
}

// ── CSV parser mínimo (maneja comillas, comas y saltos dentro de comillas) ──
function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);      // BOM
  const rows = []; let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i+1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cell); cell = ''; }
    else if (c === '\r') { /* ignora */ }
    else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
    else cell += c;
  }
  if (cell.length || row.length) { row.push(cell); rows.push(row); }
  return rows.filter(r => r.some(x => String(x).trim() !== ''));
}

// ── Normalizadores (mismo criterio que el blueprint) ───────────────────────
const norm = s => String(s == null ? '' : s).trim();
const nn = s => { const v = norm(s); return v === '' ? null : v; };            // vacío → null
const FECHAS_RARAS = [];   // fechas que no se pudieron interpretar (se reportan)
// Meses en español E inglés (la hoja mezcla ambos: "abr" y "Aug"). Clave = 3
// primeras letras en minúscula sin acento.
const MESES = { ene:1, jan:1, feb:2, mar:3, abr:4, apr:4, may:5, jun:6, jul:7,
  ago:8, aug:8, sep:9, set:9, oct:10, nov:11, dic:12, dec:12 };
// Año de 2 dígitos → pivote por el año actual: si 20yy quedara en el FUTURO
// (imposible para una fecha de nacimiento), es 19yy. Ej. '89'→1989, '05'→2005.
function _fixYear(y) {
  if (y.length <= 2) { const yy = parseInt(y,10), cur = new Date().getFullYear()%100; return String((yy <= cur ? 2000 : 1900) + yy); }
  return y.padStart(4,'0');
}
function _armaFecha(d, mo, y) {
  if (mo > 12 && d <= 12) { const t = d; d = mo; mo = t; }   // venía M/D → corrige a D/M
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return `${_fixYear(String(y))}-${String(mo).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
}
function normFecha(s) {
  const v = norm(s); if (!v) return null;
  let m = v.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;   // ISO
  // D/mmm/Y con mes en español (13/abr/85, 13-abr-1985, "13 abr 85")
  m = v.match(/^(\d{1,2})[\/\-. ]+([a-záéíóúñ]{3,})[\/\-. ]+(\d{2,4})$/i);
  if (m) { const mo = MESES[H(m[2]).slice(0,3)];
    if (mo) { const f = _armaFecha(parseInt(m[1],10), mo, m[3]); if (f) return f; } }
  // D/M/Y numérico (acepta / - .)
  m = v.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) { const f = _armaFecha(parseInt(m[1],10), parseInt(m[2],10), m[3]); if (f) return f; }
  FECHAS_RARAS.push(v);   // no se pudo: se guarda null (no aborta la carga) y se reporta
  return null;
}
const H = h => norm(h).toLowerCase()
  .replace(/[áàä]/g,'a').replace(/[éèë]/g,'e').replace(/[íìï]/g,'i')
  .replace(/[óòö]/g,'o').replace(/[úùü]/g,'u').replace(/\s+/g,' ');
function colIndex(headers, ...names) {
  const hh = headers.map(H); for (const n of names) { const i = hh.indexOf(H(n)); if (i > -1) return i; }
  // parcial
  for (const n of names) { const t = H(n); const i = hh.findIndex(x => x.indexOf(t) > -1); if (i > -1) return i; }
  return -1;
}

// Mapea una fila del CSV → registro de clinico.pacientes
function mapRow(headers, r, tenantId) {
  const g = (...names) => { const i = colIndex(headers, ...names); return i > -1 ? r[i] : ''; };
  return {
    tenant_id: tenantId,
    folio:               nn(g('ID')),
    nombre:              norm(g('Nombre Completo del Paciente','Nombre Completo','Nombre')),
    fecha_nacimiento:    normFecha(g('Fecha de Nacimiento','Nacimiento')),
    email:               nn(g('E-mail','Email','Correo')),
    telefono:            nn(g('Telefono','Teléfono','Celular')),
    origen:              nn(g('Origen')),
    canal:               nn(g('Canal')),
    medico_tratante:     nn(g('Medico Tratante','Médico Tratante')),
    pais:                nn(g('Pais','País')),
    idioma:              nn(g('Idioma')),
    observaciones:       nn(g('Observaciones / Notas','Observaciones','Notas')),
    razon_social:        nn(g('Razon Social','Razón Social')),
    rfc:                 nn(g('RFC')),
    codigo_postal:       nn(g('Codigo Postal','Código Postal','CP')),
    uso_cfdi:            nn(g('Uso CFDI','Uso de CFDI')),
    regimen_fiscal:      nn(g('Regimen Fiscal','Régimen Fiscal')),
    forma_pago_habitual: nn(g('Forma de Pago Habitual','Forma de Pago')),
    lista_precios:       nn(g('Lista de Precios','Lista')) || 'General'
  };
}

// campos que se comparan (Sheets vs Postgres)
const CAMPOS = ['folio','nombre','fecha_nacimiento','email','origen','canal','medico_tratante',
  'pais','idioma','observaciones','razon_social','rfc','codigo_postal','uso_cfdi','regimen_fiscal',
  'forma_pago_habitual','lista_precios'];

async function main() {
  // 1. tenant_id
  const t = await rest('GET', `tenants?slug=eq.${encodeURIComponent(SLUG)}&select=id`, { schema:'core' });
  if (!t || !t.length) { console.error(`No hallé el tenant '${SLUG}' en core.tenants. ¿Corriste 04_seed.sql?`); process.exit(1); }
  const tenantId = t[0].id;

  // 2. CSV → registros
  const rows = parseCSV(fs.readFileSync(csvPath, 'utf8'));
  if (rows.length < 2) { console.error('El CSV no tiene datos (¿exportaste la hoja Pacientes como CSV?).'); process.exit(1); }
  const headers = rows[0];
  const src = rows.slice(1).map(r => mapRow(headers, r, tenantId)).filter(x => x.nombre); // descarta filas sin nombre
  console.log(`CSV: ${src.length} pacientes con nombre (de ${rows.length - 1} filas).`);
  if (FECHAS_RARAS.length) {
    const u = [...new Set(FECHAS_RARAS)];
    console.warn(`⚠ ${FECHAS_RARAS.length} fecha(s) de nacimiento no reconocidas → se guardan VACÍAS (no bloquean la carga). Corrige el formato en la hoja. Ejemplos: ${u.slice(0,10).join('  ·  ')}`);
  }
  // Correos repetidos: NO se bloquean (en fertilidad las parejas comparten correo),
  // pero se avisan para que decidas si es pareja legítima o duplicado real — el
  // mismo valor que hoy da el antiduplicado del ERP, pero sin frenar la migración.
  const porCorreo = new Map();
  for (const p of src) { const e = (p.email || '').toLowerCase(); if (!e) continue; (porCorreo.get(e) || porCorreo.set(e, []).get(e)).push(p.folio || '(sin ID)'); }
  const repetidos = [...porCorreo.entries()].filter(([, fs]) => fs.length > 1);
  if (repetidos.length) {
    console.warn(`⚠ ${repetidos.length} correo(s) compartidos por >1 paciente (normal en parejas; revisa si alguno es duplicado real):`);
    repetidos.slice(0, 15).forEach(([e, fs]) => console.warn(`   · ${e}  →  ${fs.join(', ')}`));
    if (repetidos.length > 15) console.warn(`   … y ${repetidos.length - 15} más.`);
  }

  if (cmd === 'cargar') {
    // El upsert es idempotente por (tenant_id, folio). Un paciente SIN folio (ID)
    // no tiene llave: Postgres no detectaría el conflicto y lo DUPLICARÍA en cada
    // corrida, y la comparación tampoco lo podría llavear. Se excluyen y avisan.
    const sinFolio = src.filter(x => !x.folio);
    if (sinFolio.length) {
      console.error(`⚠ ${sinFolio.length} paciente(s) sin ID/folio — NO se cargan (romperían el upsert y la comparación). Ponles ID en la hoja Pacientes. Ejemplos:`);
      sinFolio.slice(0, 10).forEach(x => console.error('   · ' + x.nombre));
    }
    const cargables = src.filter(x => x.folio);
    // Prefer minimal: no pide devolver la fila (la tabla base tiene el SELECT
    // revocado para authenticated; con service_role da igual, pero es más rápido).
    let ok = 0;
    for (let i = 0; i < cargables.length; i += 200) {
      const lote = cargables.slice(i, i + 200);
      await rest('POST', 'pacientes?on_conflict=tenant_id,folio', {
        body: lote, prefer: 'resolution=merge-duplicates,return=minimal'
      });
      ok += lote.length; process.stdout.write(`\r  cargados ${ok}/${cargables.length}`);
    }
    console.log(`\n✓ Carga (upsert) terminada: ${ok} con folio${sinFolio.length ? `, ${sinFolio.length} omitidos sin ID` : ''}. Corre "verificar" para comprobar que cuadra.`);
    return;
  }

  // cmd === 'verificar'
  const dst = await restAllRows('pacientes?select=' + CAMPOS.join(',') + `&tenant_id=eq.${tenantId}&deleted_at=is.null`, 'clinico');
  console.log(`Postgres: ${dst.length} pacientes.`);
  const byFolio = new Map(dst.map(d => [String(d.folio || ''), d]));
  let faltan = 0, difs = 0, iguales = 0, sinFolio = 0; const detalle = [];
  for (const s of src) {
    if (!s.folio) { sinFolio++; continue; }   // sin llave: se avisa en 'cargar', aquí no se llavea
    const d = byFolio.get(String(s.folio));
    if (!d) { faltan++; if (detalle.length < 60) detalle.push(`  FALTA en Postgres: ${s.folio} · ${s.nombre}`); continue; }
    const camposDif = CAMPOS.filter(c => String(s[c] ?? '') !== String(d[c] ?? ''));
    if (camposDif.length) { difs++; if (detalle.length < 60) detalle.push(`  DIFIERE ${s.folio} · ${s.nombre}: ${camposDif.join(', ')}`); }
    else iguales++;
  }
  const sobran = dst.length - (iguales + difs);   // en Postgres pero no en el CSV (con folio)
  console.log(`\n== RED DE SEGURIDAD ==`);
  console.log(`  iguales:            ${iguales}`);
  console.log(`  difieren:           ${difs}`);
  console.log(`  faltan en Postgres: ${faltan}`);
  console.log(`  de más en Postgres: ${sobran > 0 ? sobran : 0}`);
  if (sinFolio) console.log(`  sin ID en el CSV:   ${sinFolio}  (no migrables hasta ponerles ID)`);
  if (detalle.length) { console.log('\nDetalle (máx 60):'); detalle.forEach(d => console.log(d)); }
  console.log('\n' + (faltan === 0 && difs === 0 && sobran <= 0 && sinFolio === 0
    ? '✓ CUADRA: Sheets y Postgres coinciden.'
    : '✗ Hay diferencias — revísalas antes de confiar en la base nueva.'));
}
main().catch(e => { console.error('Error:', e.message); process.exit(1); });
