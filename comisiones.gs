/* ═══════════════════════════════════════════════════════════════════════════
   COMISIONES POR VOLUMEN DE GRUPO (rebate a médicos externos + coordinador)
   ───────────────────────────────────────────────────────────────────────────
   LA REGLA DE NEGOCIO (textual del usuario, ya aclarada):
     "El Dr. Paladino llega con una paciente, Diana García. Ella paga su tratamiento
      directamente: Captura + Vitrificación, $23,600, precio GRUPO MEDICO. Si hacen
      entre el grupo de doctores 3 o más procedimientos en el mes les corresponde el
      10%. Ese 10% se le aplica solo al procedimiento que vaya a hacer el Dr., pero
      para llegar sí se suman todos los involucrados. Daniel Madero no hace
      procedimientos: a él se le paga una comisión del 5% de ese 10%."

   EL MODELO (4 pasos):
     1. CONTAR   — todos los procedimientos ELEGIBLES (lista configurable de
                   productos) de TODO el grupo en el mes. El grupo desbloquea.
     2. ESCALON  — la cantidad determina el nivel (progresivo: 1-2 = 0%, 3+ = 10%…).
     3. BASE     — el % del escalón se aplica SOLO a los procedimientos de ESE
                   médico. Paladino: $23,600 × 10% = $2,360. Cada quien cobra lo suyo.
     4. REPARTO  — mitad y mitad: $1,180 de crédito al Dr. Paladino (contra SU
                   cuenta) y $1,180 en efectivo a Daniel Madero. Ambos el mes
                   siguiente (`diferido`).

   POR QUE `parte` ES EL % DEL ESCALON Y NO DEL MONTO:
     El usuario dijo "5% de ese 10%". Leído literal eso sería 0.5% — no es lo que
     quiere. `parte:50` sobre un escalón de 10% da el 5% efectivo que él describe.
     Guardarlo así hace que un futuro escalón de 15% siga repartiéndose mitad y
     mitad SIN reconfigurar nada. El panel muestra el % efectivo resultante para
     que el número real quede a la vista antes de aprobar.

   FUENTE DE DATOS — UNA SOLA: BD_Ingresos + la columna OrigenExterno.
     Los dos casos del usuario aterrizan ahí:
       · La paciente paga directo  → venta a nombre de la paciente, OrigenExterno = el médico.
       · Al médico se le factura   → cargarSaldoInicial llama saveIngreso, que también
                                     crea la venta en BD_Ingresos con el médico de titular.
     (El módulo de "Descuentos por agencia" NO sirve para esto y no se toca: lee solo
      de Cuentas_Cobrar, no ve una venta pagada directo, no escribe nada y no tiene
      beneficiario — el descuento siempre baja al titular del cargo.)

   ESTE ARCHIVO NO ESCRIBE NADA POR SU CUENTA: calcularComisiones() es SOLO LECTURA.
   Solo generarComisiones() escribe, y exige permiso + aprobación explícita del
   usuario en pantalla. Es dinero, y un error se multiplica por todos los médicos.

   Requiere: origenes.gs (_origResolver), cobranza.gs (_cobTierPct, crédito a favor),
             finance.gs (saveCxP, libros por año, permisos, logAudit).
   ═══════════════════════════════════════════════════════════════════════════ */

var COMISIONES_VER        = 'comisiones-2026.07.21a';
var COMISIONES_CFG_KEY    = 'COMISIONES_CFG';
var COMISIONES_TAB        = 'Comisiones_Generadas';
var COMISIONES_PERM       = 'generar_comisiones';
// Permiso APARTE de generar: marcar un mes como "ya pagado fuera del sistema" (o
// descartarlo) NO mueve dinero, solo apaga el aviso "hay comisiones por pagar".
// Es sensible (silencia una deuda), así que se delega explícitamente por rol.
var COMISIONES_PERM_SALDAR = 'saldar_comisiones';

/* ───────────────────────── Helpers propios ─────────────────────────
   Se apoyan en cobranza.gs/finance.gs donde ya existe la pieza (_cobNum, _cobTierPct,
   _cobKeyNom, _cobKeyProd, _ingEsCancelada…). Aquí solo va lo que NO existe. */

function _comNum(v) { return (typeof _cobNum === 'function') ? _cobNum(v) : (parseFloat(String(v || '').replace(/[$,\s]/g, '')) || 0); }
function _comDeacento(s) {
  return String(s || '').replace(/[áàäâ]/g, 'a').replace(/[éèëê]/g, 'e').replace(/[íìïî]/g, 'i')
    .replace(/[óòöô]/g, 'o').replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n');
}
/* Clave de producto: _cobKeyProd (minúsculas + espacios normalizados) PERO además sin
   acentos. _cobKeyProd solo no basta aquí: si el catálogo dice "Captura + Vitrificación"
   y alguien escribe "Vitrificacion" en la config, un match exacto fallaría en silencio
   y el procedimiento NO contaría — dinero mal calculado sin un solo error visible. */
function _comKeyProd(s) {
  var base = (typeof _cobKeyProd === 'function') ? _cobKeyProd(s) : String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  return _comDeacento(base);
}
function _comKeyNom(s) {
  return (typeof _cobKeyNom === 'function') ? _cobKeyNom(s)
       : _comDeacento(String(s || '').trim().toLowerCase()).replace(/\s+/g, ' ').trim();
}
function _comPad2(n) { return ('0' + n).slice(-2); }
/* Suma n meses a 'yyyy-MM' (n negativo = hacia atrás). cobranza.gs tiene _cobMesPrev,
   pero está clavado en -1 y aquí se necesita ir HACIA ADELANTE (el pago es el mes
   siguiente, `diferido`). Esta función es su superconjunto exacto:
   _comMesSuma(m,-1) === _cobMesPrev(m) — y la prueba lo verifica, para que no diverjan. */
function _comMesSuma(mes, n) {
  var y = parseInt(String(mes).substring(0, 4), 10), m = parseInt(String(mes).substring(5, 7), 10);
  if (!y || !m) return mes;
  var t = (y * 12) + (m - 1) + (parseInt(n, 10) || 0);
  return Math.floor(t / 12) + '-' + _comPad2((t % 12) + 1);
}
function _comMesKey(anio, mes) { return parseInt(anio, 10) + '-' + _comPad2(parseInt(mes, 10)); }
/* Último día del mes 'yyyy-MM' → 'yyyy-MM-dd'. El vencimiento de la CxP: la comisión
   se paga DURANTE el mes siguiente, así que vence al cierre de ese mes. */
function _comFinDeMes(mes) {
  var y = parseInt(String(mes).substring(0, 4), 10), m = parseInt(String(mes).substring(5, 7), 10);
  return String(mes) + '-' + _comPad2(new Date(y, m, 0).getDate());
}
function _comRedondea(n) { return Math.round((_comNum(n) + Number.EPSILON) * 100) / 100; }
// Neutraliza inyección de fórmulas/CSV: un valor de usuario que empiece con = + - @
// (o tab/CR) se guardaría como FÓRMULA VIVA en la hoja (IMPORTXML, etc., ejecutándose
// bajo el dueño). Anteponer comilla lo fuerza a texto literal. Usar en TODO texto de
// usuario que se escriba a Comisiones_Generadas (hoja de dinero).
function _comCell(v) { v = String(v == null ? '' : v); return /^[=+\-@\t\r]/.test(v) ? "'" + v : v; }

/* ── Guard de deploy parcial. Apps Script comparte scope global: si origenes.gs no
   está desplegado (o está en una versión vieja), _origResolver no existe y las ventas
   NO se podrían agrupar. Reventar aquí es lo correcto: seguir sin él calcularía cero
   comisiones y parecería "este mes no hubo" en vez de "el backend está a medias". ── */
function _comGuardOrig() {
  if (typeof _origResolver !== 'function')
    return 'Actualiza origenes.gs en Apps Script y redespliega (falta _origResolver): sin la jerarquía de orígenes no se puede saber qué médico pertenece a qué grupo, y las comisiones saldrían todas en cero.';
  return '';
}
function _comGuardDeps() {
  var g = _comGuardOrig(); if (g) return g;
  if (typeof _cobTierPct !== 'function')
    return 'Actualiza cobranza.gs en Apps Script y redespliega (falta _cobTierPct): sin los escalones no se puede calcular el nivel del grupo.';
  if (typeof _cobRegistrarCreditoFavor !== 'function')
    return 'Actualiza cobranza.gs en Apps Script y redespliega (falta _cobRegistrarCreditoFavor): sin él no se puede acreditar al médico.';
  if (typeof saveCxP !== 'function' || typeof INGRESOS_IDS === 'undefined')
    return 'Actualiza finance.gs en Apps Script y redespliega (faltan saveCxP / los libros de ingresos).';
  return '';
}

/* ═════════════════════════ CONFIG (Script Property) ═════════════════════════
 * COMISIONES_CFG = { reglas:[{
 *     id, nombre, activo, grupoId,          // el grupo cuyo VOLUMEN cuenta
 *     productos:[...],                      // SOLO estos cuentan Y generan
 *     escalones:[{desde,hasta,pct}],        // progresivos; hasta:'' = sin tope
 *     beneficiarios:[{ tipo:'medico_del_procedimiento'|'fijo', origenId, parte, via }],
 *     diferido                              // meses: 1 = se aplica/paga el mes N+1
 * }] }
 * `parte` = % DEL ESCALON (suma de partes debe dar 100). `via` = 'nota_credito'|'efectivo'.
 * Sin config previa arranca vacío: el usuario captura su primera regla desde el panel.
 */
var COMISIONES_VIAS  = ['nota_credito', 'efectivo'];   // enum abierto: agregar aquí
var COMISIONES_TIPOS = ['medico_del_procedimiento', 'fijo'];

function _comCfg() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(COMISIONES_CFG_KEY);
    if (raw) { var o = JSON.parse(raw); if (o && o.reglas) return o; }
  } catch (e) {}
  return { reglas: [] };
}
function _comRegla(reglaId) {
  var found = null;
  (_comCfg().reglas || []).forEach(function (r) { if (String(r.id) === String(reglaId)) found = r; });
  return found;
}
/* Escalones normalizados para _cobTierPct: la función de cobranza compara
   `qty <= tiers[i].hasta` sin normalizar, así que un `hasta:''` la haría devolver 0
   siempre (el tope abierto quedaría cerrado). Se traduce a 1e9, igual que hace
   _cobDescuentosAgencia antes de llamarla. */
function _comTiers(regla) {
  return (regla.escalones || []).map(function (t) {
    return { desde: _comNum(t.desde) || 1,
             hasta: (t.hasta === '' || t.hasta == null) ? 1e9 : (_comNum(t.hasta) || 1e9),
             pct: _comNum(t.pct) };
  }).sort(function (a, b) { return a.desde - b.desde; });
}

/* Valida una regla. Devuelve [] si está bien, o la lista de problemas.
   La suma de `parte` DEBE dar 100: si diera menos, la clínica se quedaría con un
   pedazo del rebate que ya prometió; si diera más, pagaría de más. Ninguna de las
   dos puede pasar en silencio. */
function _comValidarRegla(r) {
  var errs = [];
  if (!r || !String(r.id || '').trim()) errs.push('La regla necesita un id.');
  if (!String(r.grupoId || '').trim()) errs.push('Falta el grupo (grupoId) cuyo volumen cuenta.');
  if (!(r.productos || []).length) errs.push('Falta la lista de procedimientos que cuentan (sin ella no contaría nada).');
  if (!(r.escalones || []).length) errs.push('Faltan los escalones (desde / hasta / %).');
  var bens = r.beneficiarios || [];
  if (!bens.length) errs.push('Faltan los beneficiarios del reparto.');
  var suma = 0;
  bens.forEach(function (b, i) {
    suma += _comNum(b.parte);
    if (COMISIONES_TIPOS.indexOf(String(b.tipo)) < 0) errs.push('Beneficiario ' + (i + 1) + ': tipo desconocido "' + b.tipo + '".');
    if (COMISIONES_VIAS.indexOf(String(b.via)) < 0)   errs.push('Beneficiario ' + (i + 1) + ': vía de pago desconocida "' + b.via + '".');
    if (String(b.tipo) === 'fijo' && !String(b.origenId || '').trim())
      errs.push('Beneficiario ' + (i + 1) + ': un beneficiario fijo necesita su origenId del catálogo.');
  });
  if (bens.length && Math.abs(suma - 100) > 0.01)
    errs.push('El reparto suma ' + suma + '% y debe sumar exactamente 100% (es el reparto DEL escalón, no del monto).');
  var tiers = _comTiers(r || {});
  for (var i = 1; i < tiers.length; i++)
    if (tiers[i].desde <= tiers[i - 1].hasta && tiers[i - 1].hasta !== 1e9)
      errs.push('Los escalones se enciman entre ' + tiers[i - 1].desde + '-' + tiers[i - 1].hasta + ' y ' + tiers[i].desde + '-' + (tiers[i].hasta === 1e9 ? '∞' : tiers[i].hasta) + '.');
  return errs;
}

/* Preview del reparto para el panel: traduce `parte` a % EFECTIVO sobre la venta.
   "Del 10% del escalón: 50% al médico = 5% · 50% a Madero = 5%" — el usuario dijo
   "5% de ese 10%", que leído literal sería 0.5%. Este preview desambigua ANTES de
   guardar la regla. */
function _comPreviewReparto(regla) {
  var tiers = _comTiers(regla || {});
  var out = [];
  tiers.forEach(function (t) {
    if (t.pct <= 0) return;
    (regla.beneficiarios || []).forEach(function (b) {
      out.push({ escalonPct: t.pct, desde: t.desde, hasta: t.hasta === 1e9 ? '' : t.hasta,
                 tipo: b.tipo, origenId: b.origenId || '', via: b.via,
                 parte: _comNum(b.parte),
                 efectivoPct: _comRedondea(t.pct * _comNum(b.parte) / 100) });
    });
  });
  return out;
}

function readComisionesCfg() {
  try {
    var cfg = _comCfg();
    var reglas = (cfg.reglas || []).map(function (r) {
      var c = JSON.parse(JSON.stringify(r));
      c._preview = _comPreviewReparto(r);
      c._errores = _comValidarRegla(r);
      return c;
    });
    return { ok: true, version: COMISIONES_VER, reglas: reglas, vias: COMISIONES_VIAS, tipos: COMISIONES_TIPOS,
             deploy: _comGuardDeps() || '' };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

function saveComisionesCfg(body) {
  try {
    body = body || {};
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token, COMISIONES_PERM))
      return { ok: false, error: 'No tienes permiso para configurar comisiones (' + COMISIONES_PERM + ').', version: COMISIONES_VER };
    var list = body.reglas || [];
    if (!Array.isArray(list)) return { ok: false, error: 'Formato inválido: se esperaba una lista de reglas.', version: COMISIONES_VER };
    // Validar TODAS antes de guardar ninguna: media config es peor que ninguna.
    var problemas = [];
    list.forEach(function (r) {
      _comValidarRegla(r).forEach(function (e) { problemas.push((r && r.nombre ? r.nombre : (r && r.id) || '(sin nombre)') + ': ' + e); });
    });
    if (problemas.length) return { ok: false, error: 'La configuración tiene errores:\n· ' + problemas.join('\n· '), problemas: problemas, version: COMISIONES_VER };
    var limpio = list.map(function (r) {
      return { id: String(r.id).trim(), nombre: String(r.nombre || '').trim(), activo: r.activo !== false,
               grupoId: String(r.grupoId).trim(),
               // Modo de cálculo: 'escalon' (base × % del escalón, reparto a beneficiarios)
               // o 'lista' (montos EXACTOS de la tabla «Tarifas por tier»: solo comisión
               // de Daniel en efectivo; el médico ya pagó descontado el precio del tier).
               modo: (String(r.modo || 'escalon') === 'lista') ? 'lista' : 'escalon',
               productos: (r.productos || []).map(function (p) { return String(p).trim(); }).filter(function (p) { return !!p; }),
               productosSoloDescuento: (r.productosSoloDescuento || []).map(function (p) { return String(p).trim(); }).filter(function (p) { return !!p; }),
               escalones: (r.escalones || []).map(function (t) { return { desde: _comNum(t.desde) || 1, hasta: (t.hasta === '' || t.hasta == null) ? '' : _comNum(t.hasta), pct: _comNum(t.pct) }; }),
               beneficiarios: (r.beneficiarios || []).map(function (b) { return { tipo: String(b.tipo), origenId: String(b.origenId || '').trim(), parte: _comNum(b.parte), via: String(b.via) }; }),
               diferido: (r.diferido == null || r.diferido === '') ? 1 : (parseInt(r.diferido, 10) || 0),
               notas: String(r.notas || '') };
    });
    PropertiesService.getScriptProperties().setProperty(COMISIONES_CFG_KEY, JSON.stringify({ reglas: limpio }));
    try { logAudit(body.usuario || '', 'Comisiones', 'Guardar configuración', '', '', '', limpio.length + ' regla(s)'); } catch (e) {}
    return { ok: true, version: COMISIONES_VER, guardadas: limpio.length };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═══════════ TARIFAS GRUPO MÉDICO — listas de precio por TIER (exactas) ═══════════
 * Reproduce EXACTO las tablas del usuario: por tratamiento×tier, el Precio a Médico
 * Externo (lo que paga el médico) + la Comisión de Daniel (efectivo). Hestia se queda
 * con precioMedico − comisionDaniel. El tier lo da el volumen del grupo en el mes.
 * Es una LISTA DE PRECIOS (lookup), no un %: da los montos exactos, no aproximados. */
var GM_TARIFAS_KEY = 'GRUPO_MEDICO_TARIFAS';
function _gmTarifasSeed(){
  // Índices 0..5 = Tier 0..5. Verificado: Percepción Hestia = precioMedico − comisionDaniel.
  return {
    tiers: [
      { n:0, label:'Tier 0', desde:1,  hasta:2,  pct:0  },
      { n:1, label:'Tier 1', desde:3,  hasta:7,  pct:10 },
      { n:2, label:'Tier 2', desde:8,  hasta:12, pct:15 },
      { n:3, label:'Tier 3', desde:13, hasta:15, pct:18 },
      { n:4, label:'Tier 4', desde:16, hasta:18, pct:22 },
      { n:5, label:'Tier 5', desde:19, hasta:'', pct:25 }
    ],
    tratamientos: [
      { nombre:'Captura + Vitrificación',                sku:'', base:23600, precio:[23600,22302,21665,21094,20249,20001], daniel:[0,1062,1605,1742,1841,2301] },
      { nombre:'Captura, Fert. y Vitrificación',         sku:'', base:40250, precio:[40250,38036,36950,35975,34535,34112], daniel:[0,1811,2737,2970,3140,3924] },
      { nombre:'Captura, Fert. y Transf. Fresco',        sku:'', base:46700, precio:[46700,44132,42871,41740,40069,39578], daniel:[0,2102,3176,3446,3643,4553] },
      { nombre:'Captura, Fert. y Transf. Diferida',      sku:'', base:53000, precio:[53000,50085,48654,47371,45474,44918], daniel:[0,2385,3604,3911,4134,5167] },
      { nombre:'Desvitrificación + Transferencia',       sku:'', base:19250, precio:[19250,18191,17672,17206,16517,16314], daniel:[0,866,1309,1421,1502,1877] },
      { nombre:'Desvit., Fert., Transf. y Vitrif.',      sku:'', base:26700, precio:[26700,25232,24511,23864,22909,22628], daniel:[0,1202,1816,1970,2083,2603] },
      { nombre:'Histeroscopia Con',                      sku:'', base:23000, precio:[23000,21850,20700,20010,19090,18400], daniel:[0,1150,1750,2070,2530,2875] },
      { nombre:'Histeroscopia Sin',                      sku:'', base:20000, precio:[20000,19000,18000,17400,16600,16000], daniel:[0,1000,1500,1800,2200,2500] },
      { nombre:'Recepción de Células',                   sku:'', base:5400,  precio:[5400,5103,4957,4827,4633,4577],       daniel:[0,243,367,399,421,527] }
    ]
  };
}
function readTarifasGM(){
  try{
    var raw = PropertiesService.getScriptProperties().getProperty(GM_TARIFAS_KEY);
    var cfg = raw ? JSON.parse(raw) : _gmTarifasSeed();
    return { ok:true, cfg:cfg, seeded:!raw };
  }catch(ex){ return { ok:false, error:ex.message, cfg:_gmTarifasSeed() }; }
}
function saveTarifasGM(body){
  try{
    if (typeof _tokenHasPermission === 'function' && !_tokenHasPermission(body.token, COMISIONES_PERM))
      return { ok:false, error:'Sin permiso para configurar comisiones ('+COMISIONES_PERM+').' };
    var cfg = body.cfg || {};
    if (!cfg.tratamientos || !cfg.tiers) return { ok:false, error:'Config incompleta (faltan tiers o tratamientos).' };
    // Sanitiza: números limpios, arreglos de 6 tiers.
    var limpio = {
      tiers: (cfg.tiers||[]).map(function(t,i){ return { n:i, label:String(t.label||('Tier '+i)), desde:_comNum(t.desde)||1, hasta:(t.hasta===''||t.hasta==null)?'':_comNum(t.hasta), pct:_comNum(t.pct) }; }),
      tratamientos: (cfg.tratamientos||[]).map(function(tr){
        function arr6(a){ a=a||[]; var o=[]; for(var k=0;k<6;k++) o.push(_comNum(a[k])||0); return o; }
        return { nombre:String(tr.nombre||'').trim(), sku:String(tr.sku||'').trim(), base:_comNum(tr.base)||0, precio:arr6(tr.precio), daniel:arr6(tr.daniel) };
      }).filter(function(tr){ return !!tr.nombre; })
    };
    PropertiesService.getScriptProperties().setProperty(GM_TARIFAS_KEY, JSON.stringify(limpio));
    try{ logAudit(body.usuario||'', 'Comisiones', 'Guardar tarifas Grupo Médico', '', '', '', limpio.tratamientos.length+' tratamiento(s)'); }catch(e){}
    return { ok:true, guardados:limpio.tratamientos.length };
  }catch(ex){ return { ok:false, error:ex.message }; }
}
// Tier (0..5) según el volumen de procedimientos del grupo en el mes.
function _gmTierPorVolumen(qty){
  var cfg = readTarifasGM().cfg; var q = Number(qty)||0;
  var ts = cfg.tiers||[];
  for (var i=0;i<ts.length;i++){ var d=Number(ts[i].desde)||1, h=(ts[i].hasta===''||ts[i].hasta==null)?1e9:Number(ts[i].hasta); if(q>=d && q<=h) return ts[i].n; }
  return 0;
}
// Lookup exacto: {precioMedico, comisionDaniel, base, descuentoMedico} para un tratamiento en un tier.
function _gmTarifa(nombreOSku, tierIdx){
  var cfg = readTarifasGM().cfg; var ti = Math.max(0, Math.min(5, Number(tierIdx)||0));
  // Match con la MISMA normalización que la elegibilidad (_comKeyProd: sin acentos +
  // minúsculas + espacios colapsados). Un lowercase pelón fallaba con "Vitrificación"
  // vs "Vitrificacion" → tf=null → comisión $0 en silencio. Ahora empata igual que el %.
  var key = _comKeyProd(nombreOSku);
  var tr = null, list = cfg.tratamientos||[];
  for (var i=0;i<list.length;i++){
    if (_comKeyProd(list[i].nombre)===key || (list[i].sku && _comKeyProd(list[i].sku)===key)) { tr=list[i]; break; }
  }
  if(!tr) return null;
  var precio = (tr.precio&&tr.precio[ti]!=null)?tr.precio[ti]:tr.base;
  var daniel = (tr.daniel&&tr.daniel[ti]!=null)?tr.daniel[ti]:0;
  return { tratamiento:tr.nombre, tier:ti, base:tr.base, precioMedico:precio, comisionDaniel:daniel,
           descuentoMedico:Math.max(0, tr.base-precio), percepcionHestia:precio-daniel };
}

/* Contexto de tarifas GM para AUTO-PRECIO en la captura de ingresos. SOLO LECTURA.
 * El usuario decidió: tier del MES ANTERIOR, estable, sin true-up → el volumen de
 * procedimientos GM del mes previo fija el tier con el que se cobra TODO el mes en curso.
 * Se recalibra solo al cambiar de mes. Devuelve la tabla ya resuelta a ese tier para que
 * el front ponga el precio del médico y muestre la comisión de Daniel (que la GENERA el
 * corte mensual, no la captura). "Volumen del grupo" = procedimientos del mes anterior
 * cuyo producto es uno de los tratamientos GM (los 9 de la tabla). */
function tarifaGMContexto(anio, mes){
  try{
    var cfg = readTarifasGM().cfg;
    var now = new Date();
    anio = parseInt(anio,10) || now.getFullYear();
    mes  = parseInt(mes,10)  || (now.getMonth()+1);
    var pm = mes-1, pa = anio; if (pm < 1){ pm = 12; pa = anio-1; }
    var mesPrevKey = pa + '-' + _comPad2(pm);
    // Llaves de los tratamientos GM (nombre y sku) para contar y para matchear en el front.
    var keys = {};
    (cfg.tratamientos||[]).forEach(function(t){
      if (t.nombre) keys[_comKeyProd(t.nombre)] = 1;
      if (t.sku)    keys[_comKeyProd(t.sku)]    = 1;
    });
    var conteo = 0;
    try{
      if (typeof INGRESOS_IDS !== 'undefined' && INGRESOS_IDS[pa]){
        var lect = _comReadVentas(pa, mesPrevKey);
        if (lect.ok) lect.rows.forEach(function(v){ if (keys[_comKeyProd(v.producto)]) conteo += (_comNum(v.cantidad)||1); });
      }
    }catch(e){}
    var tier = _gmTierPorVolumen(conteo);
    var tl = (cfg.tiers && cfg.tiers[tier]) ? cfg.tiers[tier].label : ('Tier '+tier);
    var pct = (cfg.tiers && cfg.tiers[tier]) ? cfg.tiers[tier].pct : 0;
    // Procedimientos que APLICAN según las reglas de comisión (imagen "PROCEDIMIENTOS
    // QUE APLICAN"): estos productos NO viven en la BD de productos, hay que ofrecerlos
    // en el dropdown de la cotización. Unión de todas las reglas activas de comisión.
    var procs = {};
    try{
      (_comCfg().reglas || []).forEach(function(r){
        if (r.activo === false) return;
        (r.productos || []).forEach(function(p){ if(p) procs[String(p).trim()] = 1; });
        (r.productosSoloDescuento || []).forEach(function(p){ if(p) procs[String(p).trim()] = 1; });
      });
    }catch(e){}
    return { ok:true, tier:tier, tierLabel:tl, pct:pct, conteo:conteo, mesAnterior:mesPrevKey,
             procedimientos: Object.keys(procs),
             tratamientos:(cfg.tratamientos||[]).map(function(t){
               return { nombre:t.nombre, sku:t.sku, base:t.base,
                        precioMedico:(t.precio&&t.precio[tier]!=null)?t.precio[tier]:t.base,
                        comisionDaniel:(t.daniel&&t.daniel[tier]!=null)?t.daniel[tier]:0 };
             }) };
  }catch(ex){ return { ok:false, error:ex.message }; }
}

/* ═════════════════════════ HOJA DE CONTROL ═════════════════════════
 * Comisiones_Generadas es la MEMORIA de lo ya pagado — la única defensa contra
 * generar el mismo mes dos veces. Vive en INGRESOS_SS_ID junto a Creditos_Favor.
 * Esquema: Mes | ReglaId | Beneficiario | Via | Monto | RefId | Estado | Usuario | Timestamp
 *   RefId  = OP sintética (crédito) o el ID de la fila de Egresos (CxP) → permite
 *            volver a encontrar lo escrito para revisarlo o revertirlo.
 *   Estado = 'activa' | 'revertida'
 */
function _comEnsureControl() {
  var ss = SpreadsheetApp.openById(INGRESOS_SS_ID);
  var sh = ss.getSheetByName(COMISIONES_TAB);
  if (!sh) {
    sh = ss.insertSheet(COMISIONES_TAB);
    sh.appendRow(['Mes', 'ReglaId', 'Beneficiario', 'Via', 'Monto', 'RefId', 'Estado', 'Usuario', 'Timestamp']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f3f4f6');
  }
  return sh;
}
/* Lo ya generado (dinero REAL: crédito o CxP) para un mes+regla. Vacío = nunca se
   ha generado en el sistema. SOLO cuenta estado 'activa': ni 'revertida' ni las
   marcas manuales ('saldada'/'descartada', que NO son movimientos de dinero) deben
   contar como generado, o el reversa intentaría "deshacer" una marca sin refId. */
function _comYaGenerado(mes, reglaId) {
  var out = [];
  try {
    var sh = _comEnsureControl();
    var raw = sh.getDataRange().getValues();
    for (var i = 1; i < raw.length; i++) {
      if (String(raw[i][0] || '').trim() !== String(mes)) continue;
      if (String(raw[i][1] || '').trim() !== String(reglaId)) continue;
      if (String(raw[i][6] || '').trim().toLowerCase() !== 'activa') continue;
      out.push({ rowNum: i + 1, mes: String(raw[i][0]), reglaId: String(raw[i][1]),
                 beneficiario: String(raw[i][2]), via: String(raw[i][3]),
                 monto: _comNum(raw[i][4]), refId: String(raw[i][5] || '').trim(),
                 estado: String(raw[i][6] || ''), usuario: String(raw[i][7] || ''),
                 timestamp: raw[i][8] });
    }
  } catch (e) {}
  return out;
}

/* Marca MANUAL viva de un mes+regla: "ya pagada fuera del sistema" ('saldada') o
   "no aplica" ('descartada'). NO es dinero: es un apagador del aviso de pendientes
   para los meses que se liquidaron antes de existir este módulo. Devuelve la última
   marca viva, o null. Via='marca', RefId vacío → nunca se confunde con dinero real. */
function _comMarca(mes, reglaId) {
  try {
    var sh = _comEnsureControl();
    var raw = sh.getDataRange().getValues();
    var found = null;
    for (var i = 1; i < raw.length; i++) {
      if (String(raw[i][0] || '').trim() !== String(mes)) continue;
      if (String(raw[i][1] || '').trim() !== String(reglaId)) continue;
      var est = String(raw[i][6] || '').trim().toLowerCase();
      if (est === 'saldada' || est === 'descartada')
        found = { rowNum: i + 1, tipo: est, nota: String(raw[i][2] || ''),
                  monto: _comNum(raw[i][4]), usuario: String(raw[i][7] || ''), timestamp: raw[i][8] };
    }
    return found;
  } catch (e) { return null; }
}

/* OP sintética del crédito: estable y única por (regla, mes, médico). _cobRegistrarCreditoFavor
   hace UPSERT POR OP, así que esta llave es lo que impide que regenerar apile créditos. */
function _comOpCredito(reglaId, mes, origenId) { return 'COM-' + reglaId + '-' + mes + '-' + origenId; }

/* ═════════════════════════ MOTOR (SOLO LECTURA) ═════════════════════════ */

/* Lee BD_Ingresos del año pedido. Columnas SIEMPRE por encabezado: OrigenExterno es
   una columna dinámica (_ingColEnsure la crea al final cuando falta), así que leerla
   por índice fijo es un bug esperando su turno. */
function _comReadVentas(anio, mesKey) {
  // Año sin libro → RECHAZAR. _ingIdDeAnio cae al libro 2026 cuando no conoce el año:
  // pedir 2027 devolvería datos de 2026 y calcularía comisiones sobre el mes equivocado
  // sin decir una palabra. Preferimos el error.
  if (typeof INGRESOS_IDS === 'undefined' || !INGRESOS_IDS[parseInt(anio, 10)])
    return { ok: false, error: 'No hay libro de ingresos para ' + anio + '. Los libros configurados son: ' + Object.keys(INGRESOS_IDS || {}).join(', ') + '.' };
  var ssId = _ingIdDeAnio(anio);
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName(typeof BD_INGRESOS_TAB !== 'undefined' ? BD_INGRESOS_TAB : 'BD_Ingresos');
  if (!sh) return { ok: false, error: 'No se encontró BD_Ingresos en el libro de ' + anio + '.' };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, rows: [], nombresPaciente: {} };

  var hdr = data[0].map(function (c) { return _comDeacento(String(c).trim().toLowerCase()); });
  function col(keys, fb) {
    for (var n = 0; n < keys.length; n++) for (var c = 0; c < hdr.length; c++) if (hdr[c] === keys[n]) return c;
    for (var n2 = 0; n2 < keys.length; n2++) for (var c2 = 0; c2 < hdr.length; c2++) if (hdr[c2].indexOf(keys[n2]) > -1) return c2;
    return fb;
  }
  var iOp = col(['op'], 0), iFecha = col(['fecha'], 2), iPac = col(['paciente'], 3),
      iProd = col(['producto'], 5), iCant = col(['cantidad'], 8),
      iTot = col(['totalpagar', 'total a pagar', 'total'], 9),
      iOrigen = col(['origenexterno', 'origen'], -1),
      // Sin fallback: si la hoja del año no trae la columna, nada está cancelado.
      iCancel = col(['cancelada'], -1);

  var rows = [], nombresPaciente = {};
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var op = String(r[iOp] || '').trim(); if (!op) continue;
    // Todo nombre de titular visto, para avisar si el crédito de un médico quedaría huérfano.
    if (String(r[iPac] || '').trim()) nombresPaciente[_comKeyNom(r[iPac])] = String(r[iPac]).trim();
    // VENTA CANCELADA → fuera. Una comisión sobre una venta cancelada es dinero regalado:
    // la venta se reversó en bancos pero la fila sigue en la hoja con su TotalPagar intacto.
    if (iCancel > -1 && typeof _ingEsCancelada === 'function' && _ingEsCancelada(r[iCancel])) continue;
    var d = (r[iFecha] instanceof Date) ? r[iFecha] : (typeof _cobD === 'function' ? _cobD(r[iFecha]) : new Date(r[iFecha]));
    if (!d || isNaN(d.getTime())) continue;
    var mk = d.getFullYear() + '-' + _comPad2(d.getMonth() + 1);
    if (mesKey && mk !== mesKey) continue;
    rows.push({ op: op, fecha: mk, paciente: String(r[iPac] || '').trim(),
                producto: String(r[iProd] || '').trim(),
                cantidad: _comNum(r[iCant]) || 1,
                total: _comNum(r[iTot]),
                origen: iOrigen > -1 ? String(r[iOrigen] || '').trim() : '' });
  }
  return { ok: true, rows: rows, nombresPaciente: nombresPaciente };
}

/* calcularComisiones({anio, mes, reglaId}) → SOLO LECTURA. No escribe absolutamente nada.
   Es lo que alimenta la pantalla de revisión previa: el usuario ve el detalle línea por
   línea y los totales por beneficiario ANTES de que exista un solo peso. */
function calcularComisiones(body) {
  try {
    body = body || {};
    var dep = _comGuardDeps(); if (dep) return { ok: false, error: dep, version: COMISIONES_VER };

    var anio = parseInt(body.anio, 10), mes = parseInt(body.mes, 10);
    if (!anio || !mes || mes < 1 || mes > 12) return { ok: false, error: 'Indica un año y un mes válidos.', version: COMISIONES_VER };
    var mesKey = _comMesKey(anio, mes);

    var regla = _comRegla(body.reglaId);
    if (!regla) return { ok: false, error: 'No existe la regla "' + body.reglaId + '". Configúrala en Panel de Control → Comisiones.', version: COMISIONES_VER };
    if (regla.activo === false) return { ok: false, error: 'La regla "' + (regla.nombre || regla.id) + '" está desactivada.', version: COMISIONES_VER };
    var errs = _comValidarRegla(regla);
    if (errs.length) return { ok: false, error: 'La regla tiene errores de configuración:\n· ' + errs.join('\n· '), problemas: errs, version: COMISIONES_VER };

    var lect = _comReadVentas(anio, mesKey);
    if (!lect.ok) return { ok: false, error: lect.error, version: COMISIONES_VER };

    var prodSet = {};
    (regla.productos || []).forEach(function (p) { prodSet[_comKeyProd(p)] = 1; });
    // Set B: productos que RECIBEN la comisión pero NO cuentan para el escalón/tier.
    // Elegibilidad para el beneficio = A∪B; conteo del escalón = SOLO A. Retrocompatible:
    // una regla sin este campo trata B como [] → comportamiento idéntico al anterior.
    var soloSet = {};
    (regla.productosSoloDescuento || []).forEach(function (p) { soloSet[_comKeyProd(p)] = 1; });

    // ── 1) Ventas del GRUPO con producto elegible ──────────────────────────
    var elegibles = [], descartadas = { sinOrigen: 0, otroGrupo: 0, productoNoElegible: 0 };
    // Detalle SOLO de las que SÍ son del grupo pero su PRODUCTO no está en la lista
    // elegible: es lo accionable ("Sasha con Paladino no cuenta porque Inseminación
    // no está en los productos de la regla"). sinOrigen/otroGrupo solo se cuentan.
    var descNoElegible = [];
    lect.rows.forEach(function (v) {
      if (!v.origen) { descartadas.sinOrigen++; return; }
      var res = null;
      try { res = _origResolver(v.origen); } catch (e) { res = null; }
      if (!res || !res.id) { descartadas.sinOrigen++; return; }
      if (String(res.grupoId) !== String(regla.grupoId)) { descartadas.otroGrupo++; return; }
      var pk = _comKeyProd(v.producto), enTier = !!prodSet[pk], enSolo = !!soloSet[pk];
      if (!enTier && !enSolo) {
        descartadas.productoNoElegible++;
        descNoElegible.push({ op: v.op, paciente: v.paciente, producto: v.producto,
                              cantidad: v.cantidad, total: v.total, medicoNombre: res.nombre });
        return;
      }
      elegibles.push({ op: v.op, paciente: v.paciente, producto: v.producto, cantidad: v.cantidad,
                       total: v.total, medicoId: res.id, medicoNombre: res.nombre, medicoTipo: res.tipo,
                       cuentaTier: enTier });
    });

    // ── 2) Conteo del grupo → escalón ───────────────────────────────────────
    // Se cuentan PROCEDIMIENTOS, no renglones: una venta con cantidad 2 son 2 procedimientos.
    // SOLO los productos que cuentan para tier (Set A) empujan el escalón; los "solo
    // descuento" (Set B) reciben el % alcanzado sin sumarlo.
    var conteo = 0;
    elegibles.forEach(function (e) { if (e.cuentaTier) conteo += e.cantidad; });
    var tiers = _comTiers(regla);
    var pct = _cobTierPct(tiers, conteo);
    var escalon = null;
    tiers.forEach(function (t) { if (conteo >= t.desde && conteo <= t.hasta) escalon = t; });

    // ── 3) Base POR MÉDICO (el grupo desbloquea; cada quien cobra lo suyo) ──
    var porMedico = {};
    elegibles.forEach(function (e) {
      if (!porMedico[e.medicoId]) porMedico[e.medicoId] = { medicoId: e.medicoId, medicoNombre: e.medicoNombre, base: 0, cantidad: 0, ventas: [] };
      porMedico[e.medicoId].base += e.total;
      porMedico[e.medicoId].cantidad += e.cantidad;
      porMedico[e.medicoId].ventas.push(e);
    });

    // ── 4) Reparto entre beneficiarios ──────────────────────────────────────
    var mesAplica = _comMesSuma(mesKey, (regla.diferido == null ? 1 : parseInt(regla.diferido, 10) || 0));
    var detalle = [], totales = {}, avisos = [];

    // ════════════════════ MODO LISTA DE TARIFAS ════════════════════
    // Comisión por la TABLA «Tarifas por tier» (montos EXACTOS), no por %. El médico ya
    // pagó el precio del tier al capturar (su descuento va en el precio), así que aquí
    // SOLO se genera la comisión de Daniel en efectivo: Σ comisionDaniel[tier] × cantidad
    // de cada procedimiento del grupo. Los beneficiarios en nota de crédito se ignoran.
    if (String(regla.modo || 'escalon') === 'lista') {
      // Tier con la MISMA fuente que fija el PRECIO al capturar (_gmTierPorVolumen lee
      // los rangos desde/hasta de la propia tabla de tarifas y devuelve el n=0..5 que
      // indexa daniel[]/precio[]). NO se usa la posición del escalón de la regla: esa es
      // otra config y podría no coincidir → pagaría a Daniel en un tier y cobraría al
      // médico en otro. Así ambos (precio y comisión) salen del mismo tier.
      var tierIdx = _gmTierPorVolumen(conteo);
      var benEf = (regla.beneficiarios || []).filter(function (b) { return String(b.via) === 'efectivo'; })[0];
      var danId = '', danNombre = '';
      if (benEf && String(benEf.tipo) === 'fijo') {
        var rfD = null; try { rfD = _origResolver(benEf.origenId); } catch (e) {}
        if (rfD && rfD.id) { danId = rfD.id; danNombre = rfD.nombre; }
      }
      var avisosL = [];
      if (!benEf) avisosL.push('No hay beneficiario en EFECTIVO (Daniel) en el reparto de la regla: no se generará nada. Agrégalo.');
      else if (!danId) avisosL.push('El beneficiario en efectivo no se pudo resolver en el catálogo de orígenes (revisa su origenId).');
      var totalDaniel = 0, lineasD = [], sinTarifa = [];
      elegibles.forEach(function (e) {
        var tf = null; try { tf = _gmTarifa(e.producto, tierIdx); } catch (ex) { tf = null; }
        var unit = tf ? _comNum(tf.comisionDaniel) : 0;
        var monto = _comRedondea(unit * e.cantidad);
        if (!tf || unit <= 0) sinTarifa.push(e.producto);
        if (monto > 0.01) {
          totalDaniel = _comRedondea(totalDaniel + monto);
          lineasD.push({ op: e.op, paciente: e.paciente, producto: e.producto, cantidad: e.cantidad,
                         medicoNombre: e.medicoNombre, comisionUnit: unit, monto: monto, tier: tierIdx });
        }
      });
      if (sinTarifa.length) {
        var uniqST = sinTarifa.filter(function (v, i, a) { return a.indexOf(v) === i; });
        avisosL.push('Sin comisión Daniel en la tabla (tier ' + tierIdx + ') para: ' + uniqST.join(', ') + '. Llénalas en «💵 Tarifas por tier» o no generan comisión.');
      }
      var totalesL = [];
      if (danId && totalDaniel > 0.01) totalesL.push({ beneficiarioId: danId, beneficiario: danNombre, via: 'efectivo', monto: totalDaniel, mesAplica: mesAplica, lineas: lineasD });
      var yaL = _comYaGenerado(mesKey, regla.id);
      return { ok: true, version: COMISIONES_VER, modo: 'lista',
               mes: mesKey, mesAplica: mesAplica, anio: anio,
               regla: { id: regla.id, nombre: regla.nombre, diferido: regla.diferido == null ? 1 : regla.diferido, productos: regla.productos, modo: 'lista' },
               conteo: conteo, tier: tierIdx,
               escalon: escalon ? { desde: escalon.desde, hasta: escalon.hasta === 1e9 ? '' : escalon.hasta, pct: escalon.pct } : null,
               pct: pct, escalones: tiers.map(function (t) { return { desde: t.desde, hasta: t.hasta === 1e9 ? '' : t.hasta, pct: t.pct }; }),
               detalle: lineasD, totales: totalesL, totalGeneral: totalDaniel,
               elegibles: elegibles, descartadas: descartadas, descartadasDetalle: descNoElegible,
               avisos: avisosL, bloqueo: '',
               yaGenerado: yaL, generado: yaL.length > 0, marca: _comMarca(mesKey, regla.id) };
    }

    if (pct > 0) {
      Object.keys(porMedico).forEach(function (mid) {
        var M = porMedico[mid];
        var accrual = _comRedondea(M.base * pct / 100);
        (regla.beneficiarios || []).forEach(function (b) {
          var monto = _comRedondea(accrual * _comNum(b.parte) / 100);
          if (monto <= 0.01) return;
          var benId = '', benNombre = '';
          if (String(b.tipo) === 'medico_del_procedimiento') { benId = M.medicoId; benNombre = M.medicoNombre; }
          else {
            var rf = null; try { rf = _origResolver(b.origenId); } catch (e) { rf = null; }
            if (!rf || !rf.id) { avisos.push('El beneficiario fijo "' + b.origenId + '" no existe en el catálogo de orígenes: su parte no se generó.'); return; }
            benId = rf.id; benNombre = rf.nombre;
          }
          detalle.push({ medicoId: M.medicoId, medicoNombre: M.medicoNombre,
                         base: _comRedondea(M.base), cantidadMedico: M.cantidad,
                         pct: pct, accrual: accrual,
                         beneficiarioId: benId, beneficiario: benNombre,
                         tipo: b.tipo, parte: _comNum(b.parte),
                         efectivoPct: _comRedondea(pct * _comNum(b.parte) / 100),
                         via: b.via, monto: monto, mesAplica: mesAplica,
                         ventas: M.ventas.map(function (v) { return { op: v.op, paciente: v.paciente, producto: v.producto, cantidad: v.cantidad, total: v.total }; }) });
          // TOTALES POR BENEFICIARIO = lo que de verdad se escribe. Un beneficiario fijo
          // (el coordinador) cobra por los procedimientos de VARIOS médicos: si se
          // escribiera una CxP por cada médico, Madero tendría 3 pagos sueltos el mismo
          // mes en vez del pago único que describe la regla. El detalle de arriba es para
          // que se vea de dónde salió cada peso; esto es lo que se paga.
          var tk = benId + '|' + b.via;
          if (!totales[tk]) totales[tk] = { beneficiarioId: benId, beneficiario: benNombre, via: b.via, monto: 0, mesAplica: mesAplica, lineas: [] };
          totales[tk].monto = _comRedondea(totales[tk].monto + monto);
          totales[tk].lineas.push({ medicoNombre: M.medicoNombre, base: _comRedondea(M.base), pct: pct, parte: _comNum(b.parte), efectivoPct: _comRedondea(pct * _comNum(b.parte) / 100), monto: monto });
        });
      });
    }

    // ── Avisos honestos sobre las vías de pago ──────────────────────────────
    // (a) El crédito solo se puede consumir contra una cuenta cuyo titular normalizado
    //     sea IDÉNTICO. Si el médico nunca ha sido titular de una venta, el crédito
    //     nace y se queda ahí. Mejor avisarlo antes de crearlo que después.
    detalle.forEach(function (d) {
      if (d.via !== 'nota_credito') return;
      if (!lect.nombresPaciente[_comKeyNom(d.beneficiario)])
        avisos.push('El crédito de ' + d.beneficiario + ' ($' + d.monto.toLocaleString('es-MX') + ') quedará a su nombre, pero en ' + anio + ' no aparece ninguna cuenta a nombre de "' + d.beneficiario + '". Solo se podrá aplicar cuando se le facture con ESE mismo nombre.');
    });
    // (b) saveCxP escribe SIEMPRE en el libro Egresos2026 (está clavado ahí, no es
    //     year-aware). Si la comisión se paga en otro año, la CxP caería en el libro
    //     equivocado. No se genera a ciegas: se avisa aquí y se BLOQUEA al generar.
    var anioPago = parseInt(String(mesAplica).substring(0, 4), 10);
    var bloqueoLibro = '';
    var hayEfectivo = detalle.some(function (d) { return d.via === 'efectivo'; });
    if (hayEfectivo && typeof EGRESOS_IDS !== 'undefined' && typeof EGRESOS_SS_2026 !== 'undefined' && EGRESOS_IDS[anioPago] !== EGRESOS_SS_2026) {
      bloqueoLibro = 'La comisión se paga en ' + mesAplica + ', pero las cuentas por pagar se escriben siempre en el libro Egresos2026. Generarla dejaría el pago a ' + (detalle.filter(function (d) { return d.via === 'efectivo'; })[0] || {}).beneficiario + ' en el libro del año equivocado. Antes de generar: agrega el libro de ' + anioPago + ' y haz saveCxP year-aware.';
      avisos.push(bloqueoLibro);
    }

    var ya = _comYaGenerado(mesKey, regla.id);
    var totLista = Object.keys(totales).map(function (k) { return totales[k]; })
      .sort(function (a, b) { return b.monto - a.monto; });

    return { ok: true, version: COMISIONES_VER,
             mes: mesKey, mesAplica: mesAplica, anio: anio,
             regla: { id: regla.id, nombre: regla.nombre, diferido: regla.diferido == null ? 1 : regla.diferido, productos: regla.productos },
             conteo: conteo, escalon: escalon ? { desde: escalon.desde, hasta: escalon.hasta === 1e9 ? '' : escalon.hasta, pct: escalon.pct } : null,
             pct: pct, escalones: tiers.map(function (t) { return { desde: t.desde, hasta: t.hasta === 1e9 ? '' : t.hasta, pct: t.pct }; }),
             detalle: detalle, totales: totLista,
             totalGeneral: _comRedondea(totLista.reduce(function (s, t) { return s + t.monto; }, 0)),
             elegibles: elegibles, descartadas: descartadas, descartadasDetalle: descNoElegible,
             avisos: avisos, bloqueo: bloqueoLibro,
             yaGenerado: ya, generado: ya.length > 0, marca: _comMarca(mesKey, regla.id) };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═══════════════ AVISO DE DESCUENTO/REBATE POR ORIGEN (SOLO LECTURA) ═══════════════
 * Para el aviso al CAPTURAR un ingreso: dado el origen elegido (médico externo o
 * agencia) y el mes, devuelve el % que aplica ESTE MES. NO escribe nada.
 *  · Médico externo → busca su grupo (_origResolver) → la regla de comisión activa
 *    de ese grupo → calcularComisiones → pct del escalón alcanzado por el volumen
 *    del grupo. (Es un REBATE al médico, no una rebaja al precio del paciente.)
 *  · Agencia (Reprovida) → _cobDescuentosAgencia → % de volumen del mes.
 * Devuelve {ok, tipo:'medico'|'agencia'|'', nombre, pct, conteo, escalones, ...}.
 */
function descuentoOrigen(origen, anio, mes) {
  try {
    origen = String(origen || '').trim();
    if (!origen) return { ok: true, tipo: '', pct: 0 };
    var now = new Date();
    anio = parseInt(anio, 10) || now.getFullYear();
    mes  = parseInt(mes, 10)  || (now.getMonth() + 1);

    // (A) MÉDICO EXTERNO → regla de su grupo → pct del mes.
    var res = null;
    try { if (typeof _origResolver === 'function') res = _origResolver(origen); } catch (e) {}
    if (res && res.id) {
      var grupoId = String(res.grupoId || res.id);
      var reglas = (_comCfg().reglas || []).filter(function (r) {
        return r.activo !== false && String(r.grupoId) === grupoId;
      });
      if (reglas.length) {
        // El descuento que APLICA a una venta de (mes) se GANÓ 'diferido' meses antes:
        // con diferido=1, una venta de julio usa el volumen acumulado de JUNIO (el mes
        // que ya cerró). Con diferido=0 usa el mes en curso. Por eso se calcula sobre el
        // mes GANADO = mes − diferido, no sobre el mes de la venta.
        var difer = (reglas[0].diferido == null) ? 1 : (parseInt(reglas[0].diferido, 10) || 0);
        var earnKey = _comMesSuma(_comMesKey(anio, mes), -difer);
        var eAnio = parseInt(String(earnKey).substring(0, 4), 10), eMes = parseInt(String(earnKey).substring(5, 7), 10);
        var calc = calcularComisiones({ anio: eAnio, mes: eMes, reglaId: reglas[0].id });
        if (calc && calc.ok) {
          return { ok: true, tipo: 'medico', nombre: res.nombre || origen,
                   regla: reglas[0].nombre || reglas[0].id,
                   pct: calc.pct || 0, conteo: calc.conteo || 0,
                   escalon: calc.escalon || null, escalones: calc.escalones || [],
                   diferido: difer, mesGanado: calc.mes, mesAplica: _comMesKey(anio, mes) };
        }
      }
    }
    // (B) AGENCIA → descuento por volumen del mes.
    try {
      if (typeof _cobDescuentosAgencia === 'function') {
        var desc = _cobDescuentosAgencia(origen);
        if (desc && desc.aplica) {
          var pctMax = 0;
          (desc.lineas || []).forEach(function (l) { if ((_comNum(l.pct) || 0) > pctMax) pctMax = _comNum(l.pct); });
          return { ok: true, tipo: 'agencia', nombre: origen, pct: pctMax, total: desc.total || 0 };
        }
      }
    } catch (e2) {}
    return { ok: true, tipo: '', pct: 0, nombre: origen };
  } catch (ex) { return { ok: false, error: ex.message }; }
}

/* ═══════════════ COMISIONES PENDIENTES DE GENERAR (SOLO LECTURA) ═══════════════
 * Recorre las reglas activas × los últimos meses cerrados y lista las que tienen
 * pct>0 y AÚN no se han generado. Alimenta el aviso "hay comisiones por pagar".
 * Solo lectura (calcularComisiones no escribe). `meses` = cuántos meses hacia atrás
 * revisar (default 3, incluyendo el mes anterior; el mes en curso no se cierra aún).
 */
function comisionesPendientes(meses) {
  try {
    meses = parseInt(meses, 10) || 3;
    var reglas = (_comCfg().reglas || []).filter(function (r) { return r.activo !== false; });
    if (!reglas.length) return { ok: true, pendientes: [], total: 0 };
    var now = new Date();
    var pend = [], totalMonto = 0;
    // Del mes anterior hacia atrás (el mes actual sigue acumulando, no se genera).
    for (var k = 1; k <= meses; k++) {
      var d = new Date(now.getFullYear(), now.getMonth() - k, 1);
      var anio = d.getFullYear(), mes = d.getMonth() + 1;
      for (var i = 0; i < reglas.length; i++) {
        var calc = null;
        try { calc = calcularComisiones({ anio: anio, mes: mes, reglaId: reglas[i].id }); } catch (e) { calc = null; }
        if (!calc || !calc.ok) continue;
        // Pendiente = alcanzó escalón, NO generado en el sistema y NO marcado a mano.
        // Pero la marca guarda el MONTO al momento de marcar: si al mes le entraron
        // MÁS ventas después (captura tardía / back-fill de OrigenExterno), la parte
        // NUEVA vuelve a ser pendiente. Comparamos el total vivo con lo marcado y, si
        // creció, se re-avisa SOLO el excedente (no todo el mes ya liquidado).
        var marcaMonto = calc.marca ? _comNum(calc.marca.monto) : 0;
        var neto = _comRedondea((calc.totalGeneral || 0) - marcaMonto);
        if ((calc.pct || 0) > 0 && !calc.generado && (calc.totalGeneral || 0) > 0.01 && (!calc.marca || neto > 0.01)) {
          pend.push({ reglaId: reglas[i].id, regla: reglas[i].nombre || reglas[i].id,
                      anio: anio, mes: mes, mesKey: calc.mes, pct: calc.pct,
                      conteo: calc.conteo,
                      total: (calc.marca ? neto : (calc.totalGeneral || 0)),
                      crecioTrasMarcar: !!calc.marca,
                      marcadoPrevio: (calc.marca ? _comRedondea(marcaMonto) : 0),
                      totalVivo: _comRedondea(calc.totalGeneral || 0) });
          totalMonto += (calc.marca ? neto : (calc.totalGeneral || 0));
        }
      }
    }
    return { ok: true, pendientes: pend, total: _comRedondea(totalMonto) };
  } catch (ex) { return { ok: false, error: ex.message, pendientes: [] }; }
}

/* ═══════════ MARCAR COMO PAGADA / DESCARTAR (apaga el aviso, NO mueve dinero) ═══════════
 * Para los meses que se PAGARON fuera del sistema (antes de existir este módulo) o
 * que NO aplican: se escribe una MARCA en el control (Via='marca', RefId vacío,
 * Estado 'saldada'|'descartada') que solo quita ese mes del aviso "hay comisiones por
 * pagar". No crea créditos ni CxP. Gated por 'saldar_comisiones' (distinto de generar).
 *   accion: 'saldar' (pagada fuera) | 'descartar' (no aplica) | 'desmarcar' (revivir)
 */
function saldarComision(body, actorEmail) {
  try {
    body = body || {};
    if (typeof _tokenHasPermission !== 'function')
      return { ok: false, error: 'Actualiza finance.gs en Apps Script y redespliega (falta el verificador de permisos).', version: COMISIONES_VER };
    if (!_tokenHasPermission(body.token, COMISIONES_PERM_SALDAR))
      return { ok: false, error: 'No tienes permiso para marcar comisiones como pagadas (' + COMISIONES_PERM_SALDAR + '). Pídeselo al administrador o al director.', version: COMISIONES_VER };
    // El actor es la identidad VERIFICADA del token (la pasa doPost), no el cliente:
    // esta marca silencia una deuda, su autoría no puede ser falsificable.
    var actor = String(actorEmail || '').trim() || 'sistema';

    var accion = String(body.accion || 'saldar').toLowerCase();
    if (['saldar', 'descartar', 'desmarcar'].indexOf(accion) < 0)
      return { ok: false, error: 'Acción no válida: ' + accion + ' (usa saldar, descartar o desmarcar).', version: COMISIONES_VER };

    // Recalcular SIEMPRE contra la hoja: mesKey/reglaId/monto salen del cálculo real,
    // nunca de lo que mande el cliente.
    var calc = calcularComisiones({ anio: body.anio, mes: body.mes, reglaId: body.reglaId });
    if (!calc.ok) return calc;
    var mesKey = calc.mes, reglaId = calc.regla.id;

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado, reintenta en un momento.', version: COMISIONES_VER };
    try {
      var sh = _comEnsureControl();
      var marca = _comMarca(mesKey, reglaId);

      if (accion === 'desmarcar') {
        if (!marca) return { ok: false, error: 'No hay ninguna marca de ' + mesKey + ' (regla ' + reglaId + ') que quitar.', version: COMISIONES_VER };
        sh.getRange(marca.rowNum, 7).setValue('desmarcada');
        try { logAudit(actor, 'Comisiones', 'Desmarcar (reactivar pendiente)', mesKey + '/' + reglaId, marca.tipo, '', ''); } catch (e) {}
        return { ok: true, desmarcado: true, version: COMISIONES_VER, mes: mesKey, reglaId: reglaId };
      }

      // saldar / descartar: NO se puede marcar como pagado algo que el sistema YA
      // generó (créditos/CxP vivos). Para eso está Revertir, no la marca.
      var ya = _comYaGenerado(mesKey, reglaId);
      if (ya.length)
        return { ok: false, version: COMISIONES_VER,
                 error: 'Las comisiones de ' + mesKey + ' (regla ' + reglaId + ') YA están GENERADAS en el sistema (créditos/CxP). Marcarlas como pagadas las duplicaría en los registros. Si de verdad quieres deshacerlas, usa «Revertir».' };
      if (marca)
        return { ok: true, version: COMISIONES_VER, yaMarcada: true, mes: mesKey, reglaId: reglaId, tipo: marca.tipo,
                 msg: 'Ese mes ya estaba marcado como ' + (marca.tipo === 'descartada' ? 'descartado' : 'pagado fuera') + ' por ' + (marca.usuario || '—') + '.' };

      var estado = (accion === 'descartar') ? 'descartada' : 'saldada';
      var nota = String(body.nota || '').slice(0, 300) ||
                 (estado === 'saldada' ? 'Pagada fuera del sistema (antes del módulo de comisiones)' : 'Descartada / no aplica');
      // _comCell neutraliza fórmulas en la nota (texto de usuario); el actor ya es
      // el email verificado, pero se sanea igual por si acaso.
      sh.appendRow([mesKey, reglaId, _comCell(nota), 'marca', calc.totalGeneral || 0, '', estado, _comCell(actor), new Date()]);
      try {
        logAudit(actor, 'Comisiones',
                 estado === 'saldada' ? 'Marcar como pagada (externa)' : 'Descartar',
                 mesKey + '/' + reglaId, '', '$' + (calc.totalGeneral || 0), nota);
      } catch (e) {}
      return { ok: true, version: COMISIONES_VER, mes: mesKey, reglaId: reglaId, tipo: estado, monto: _comRedondea(calc.totalGeneral || 0) };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═══════════════ REPORTE GENERAL POR RANGO DE MESES (SOLO LECTURA) ═══════════════
 * Recorre [desde..hasta] (YYYY-MM) × reglas (o una) y devuelve una fila por mes+regla
 * con conteo, escalón, total y ESTADO (pendiente/generado/saldada/descartada) + los
 * beneficiarios. Alimenta el "reporte general" imprimible. No escribe nada.
 * Ojo escalabilidad: calcularComisiones lee BD_Ingresos del año en cada llamada; el
 * rango se topa a 24 meses para no exceder el tiempo de Apps Script.
 */
function reporteComisionesRango(body) {
  try {
    body = body || {};
    var desde = String(body.desde || '').trim(), hasta = String(body.hasta || '').trim();
    if (!/^\d{4}-\d{2}$/.test(desde) || !/^\d{4}-\d{2}$/.test(hasta))
      return { ok: false, error: 'Rango inválido: usa AAAA-MM en «desde» y «hasta».', version: COMISIONES_VER };
    if (desde > hasta) { var _t = desde; desde = hasta; hasta = _t; }
    var soloRegla = String(body.reglaId || '').trim();
    var reglas = (_comCfg().reglas || []);
    if (soloRegla) reglas = reglas.filter(function (r) { return r.id === soloRegla; });
    else reglas = reglas.filter(function (r) { return r.activo !== false; });
    if (!reglas.length) return { ok: true, version: COMISIONES_VER, desde: desde, hasta: hasta, filas: [], totales: {} };

    var y0 = parseInt(desde.substring(0, 4), 10), m0 = parseInt(desde.substring(5, 7), 10);
    var y1 = parseInt(hasta.substring(0, 4), 10), m1 = parseInt(hasta.substring(5, 7), 10);
    // Rechazar rangos largos ANTES de trabajar: truncar en silencio un reporte de
    // dinero (dejando fuera justo los meses recientes) sería peor que negarse.
    var totalMeses = (y1 - y0) * 12 + (m1 - m0) + 1;
    if (totalMeses > 24)
      return { ok: false, version: COMISIONES_VER,
               error: 'El rango pedido es de ' + totalMeses + ' meses; el máximo es 24 por reporte. Acórtalo (por ejemplo, un año a la vez).' };
    var filas = [], tot = { pendiente: 0, generado: 0, saldada: 0, descartada: 0, general: 0 };
    var guard = 0;
    for (var y = y0, m = m0; (y < y1) || (y === y1 && m <= m1); ) {
      for (var i = 0; i < reglas.length; i++) {
        var calc = null;
        try { calc = calcularComisiones({ anio: y, mes: m, reglaId: reglas[i].id }); } catch (e) { calc = null; }
        if (calc && calc.ok) {
          var totGen = _comRedondea(calc.totalGeneral || 0);
          // 'sin_comision' cuando el grupo NO llegó al escalón (pct=0, $0): no es
          // pendiente ni requiere acción; 'pendiente' solo si de verdad alcanzó y debe.
          var estado = 'sin_comision';
          if (calc.generado) estado = 'generado';
          else if (calc.marca) estado = (calc.marca.tipo === 'descartada' ? 'descartada' : 'saldada');
          else if ((calc.pct || 0) > 0 && totGen > 0.01) estado = 'pendiente';
          filas.push({ mesKey: calc.mes, reglaId: reglas[i].id, regla: reglas[i].nombre || reglas[i].id,
                       conteo: calc.conteo, pct: calc.pct, total: totGen, estado: estado,
                       mesAplica: calc.mesAplica,
                       beneficiarios: (calc.totales || []).map(function (t) { return { beneficiario: t.beneficiario, via: t.via, monto: _comRedondea(t.monto), aplica: t.mesAplica }; }),
                       marca: calc.marca ? { tipo: calc.marca.tipo, usuario: calc.marca.usuario, nota: calc.marca.nota } : null });
          if (totGen > 0.01) { tot.general += totGen; if (tot[estado] != null) tot[estado] += totGen; }
        }
      }
      m++; if (m > 12) { m = 1; y++; }
      if (++guard > 24) break;   // tope de seguridad: 24 meses
    }
    return { ok: true, version: COMISIONES_VER, desde: desde, hasta: hasta, filas: filas,
             totales: { general: _comRedondea(tot.general), pendiente: _comRedondea(tot.pendiente),
                        generado: _comRedondea(tot.generado), saldada: _comRedondea(tot.saldada),
                        descartada: _comRedondea(tot.descartada) } };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═════════════════════════ REVERSA (para regenerar) ═════════════════════════
 * Solo se puede revertir lo que NADIE ha tocado todavía. Se revisa TODO antes de
 * revertir NADA: una reversa a medias dejaría al médico sin crédito y a Madero con
 * su CxP viva, o al revés.
 *
 * Por qué es indispensable revisar:
 *   · Creditos_Favor.MontoCredito es el SALDO VIVO — consumirlo lo decrementa. Y
 *     _cobRegistrarCreditoFavor hace upsert ABSOLUTO. Regenerar a ciegas sobre un
 *     crédito ya gastado a medias lo REGRESARÍA a su monto original: le devolvería
 *     al médico un dinero que ya se gastó. Dinero duplicado, en silencio.
 *   · La CxP en Egresos puede estar YA PAGADA: revertirla borraría el registro de
 *     un pago que de verdad salió de la caja.
 */
function _comEstadoRef(item) {
  // → { tocado:bool, motivo:'', rowNum } para un renglón ya generado.
  try {
    if (item.via === 'nota_credito') {
      var sh = _cobEnsureCreditos();
      var raw = sh.getDataRange().getValues();
      for (var i = 1; i < raw.length; i++) {
        if (String(raw[i][1] || '').trim() !== item.refId) continue;
        var vivo = _comNum(raw[i][3]);
        // El crédito nació con item.monto. Si hoy vale menos, ya se consumió algo.
        if (Math.abs(vivo - item.monto) > 0.01)
          return { tocado: true, rowNum: i + 1,
                   motivo: 'El crédito de ' + item.beneficiario + ' ya se usó (nació en $' + item.monto.toLocaleString('es-MX') + ' y hoy quedan $' + vivo.toLocaleString('es-MX') + '). Regenerarlo le devolvería un dinero que ya se aplicó.' };
        return { tocado: false, rowNum: i + 1 };
      }
      // Ya no está: alguien lo borró a mano. No lo damos por bueno.
      return { tocado: true, rowNum: -1, motivo: 'No encuentro el crédito ' + item.refId + ' de ' + item.beneficiario + ' en Creditos_Favor (¿se borró a mano?). Revísalo antes de regenerar.' };
    }
    if (item.via === 'efectivo') {
      var ssE = SpreadsheetApp.openById(EGRESOS_SS_2026);
      var shE = ssE.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
      if (!shE) return { tocado: true, rowNum: -1, motivo: 'No se encontró la hoja de Egresos.' };
      var rawE = shE.getDataRange().getValues();
      for (var j = 1; j < rawE.length; j++) {
        if (String(rawE[j][0] || '').trim() !== String(item.refId)) continue;
        var pagado = rawE[j][13] === true || String(rawE[j][13]).toUpperCase() === 'TRUE';
        var fecha = rawE[j][1];
        if (pagado || fecha)
          return { tocado: true, rowNum: j + 1,
                   motivo: 'La comisión en efectivo de ' + item.beneficiario + ' ($' + item.monto.toLocaleString('es-MX') + ') YA SE PAGÓ. Borrarla eliminaría el registro de un pago real.' };
        return { tocado: false, rowNum: j + 1 };
      }
      return { tocado: true, rowNum: -1, motivo: 'No encuentro la cuenta por pagar ' + item.refId + ' de ' + item.beneficiario + ' en Egresos (¿se borró a mano?). Revísala antes de regenerar.' };
    }
    return { tocado: true, rowNum: -1, motivo: 'Vía desconocida: ' + item.via };
  } catch (ex) { return { tocado: true, rowNum: -1, motivo: 'No se pudo revisar ' + item.refId + ': ' + ex.message }; }
}

/* ═════════════════════════ GENERAR (el que escribe) ═════════════════════════ */
function generarComisiones(body) {
  try {
    body = body || {};
    var dep = _comGuardDeps(); if (dep) return { ok: false, error: dep, version: COMISIONES_VER };

    // PERMISO ANTES QUE NADA: ni una lectura de más si no tiene derecho a escribir.
    // Al ser un permiso nuevo, ningún rol lo trae todavía → solo admin/director
    // (que _tokenHasPermission deja pasar). Falla del lado seguro a propósito.
    if (typeof _tokenHasPermission !== 'function')
      return { ok: false, error: 'Actualiza finance.gs en Apps Script y redespliega (falta el verificador de permisos).', version: COMISIONES_VER };
    if (!_tokenHasPermission(body.token, COMISIONES_PERM))
      return { ok: false, error: 'No tienes permiso para generar comisiones (' + COMISIONES_PERM + '). Pídeselo al administrador o al director.', version: COMISIONES_VER };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado, reintenta en un momento.', version: COMISIONES_VER };
    try {
      // Recalcular AQUÍ DENTRO: nunca confiar en los montos que mande el cliente.
      // Lo que se aprobó en pantalla se vuelve a calcular contra la hoja antes de escribir.
      var calc = calcularComisiones({ anio: body.anio, mes: body.mes, reglaId: body.reglaId });
      if (!calc.ok) return calc;
      if (calc.bloqueo) return { ok: false, error: calc.bloqueo, version: COMISIONES_VER };
      if (!calc.detalle.length)
        return { ok: false, error: 'No hay nada que generar: el grupo hizo ' + calc.conteo + ' procedimiento(s) y le corresponde ' + calc.pct + '%.', version: COMISIONES_VER, calculo: calc };

      var mesKey = calc.mes, reglaId = calc.regla.id;
      var ya = _comYaGenerado(mesKey, reglaId);

      // ── SOLO REVERTIR (deshacer una generación hecha por error) ──────────
      // Misma seguridad que regenerar: se revisa TODO antes de tocar NADA; si una
      // pieza ya se cobró/usó, NO se revierte nada (para no borrar dinero real).
      if (body.soloRevertir) {
        if (!ya.length) return { ok: false, error: 'No hay comisiones generadas de ' + mesKey + ' (regla ' + reglaId + ') para revertir.', version: COMISIONES_VER };
        var estR = ya.map(function (i) { return { item: i, est: _comEstadoRef(i) }; });
        var tocR = estR.filter(function (e) { return e.est.tocado; });
        if (tocR.length)
          return { ok: false, version: COMISIONES_VER, noSePuedeRevertir: true,
                   error: 'No se puede revertir ' + mesKey + ': parte de lo generado ya se cobró/usó.\n· ' +
                          tocR.map(function (t) { return t.est.motivo; }).join('\n· ') + '\nNo se modificó nada.',
                   motivos: tocR.map(function (t) { return t.est.motivo; }) };
        var revR = [];
        estR.forEach(function (e) {
          var i = e.item;
          if (i.via === 'nota_credito') { _cobRegistrarCreditoFavor(i.refId, i.beneficiario, 0, null); }
          else if (i.via === 'efectivo' && e.est.rowNum > 1) {
            var ssDr = SpreadsheetApp.openById(EGRESOS_SS_2026);
            var shDr = ssDr.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
            shDr.deleteRow(e.est.rowNum);
            try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (_c) {}
          }
          revR.push({ beneficiario: i.beneficiario, via: i.via, monto: i.monto, refId: i.refId });
        });
        var shCr = _comEnsureControl();
        ya.map(function (i) { return i.rowNum; }).sort(function (a, b) { return b - a; })
          .forEach(function (rn) { shCr.getRange(rn, 7).setValue('revertida'); });
        try { logAudit(body.usuario || 'sistema', 'Comisiones', 'Revertir (deshacer)', mesKey + '/' + reglaId, '', '', revR.length + ' movimiento(s)'); } catch (e) {}
        return { ok: true, revertido: true, version: COMISIONES_VER, mes: mesKey, reglaId: reglaId, revertidos: revR };
      }

      // ── MES MARCADO A MANO (pagado fuera / descartado) ──────────────────
      // La autoridad del DINERO (no solo la UI) bloquea generar un mes que ya se
      // marcó como pagado fuera del sistema: generarlo crearía créditos/CxP que
      // duplicarían un pago ya hecho. Va DESPUÉS de soloRevertir (revertir dinero
      // real sigue permitido) y ANTES de escribir. Para generarlo: desmarcar primero.
      if (calc.marca) {
        return { ok: false, version: COMISIONES_VER, hayMarca: true,
                 error: 'Este mes está marcado como ' + (calc.marca.tipo === 'descartada' ? 'DESCARTADO (no aplica)' : 'YA PAGADO fuera del sistema') +
                        (calc.marca.usuario ? ' por ' + calc.marca.usuario : '') +
                        '. Generar crearía créditos/CxP y duplicaría un pago ya hecho. Si de verdad quieres generarlo, primero quita la marca (Calcular ese mes → ↩ Quitar la marca).' };
      }

      // ── IDEMPOTENCIA ────────────────────────────────────────────────────
      if (ya.length && !body.regenerar) {
        return { ok: false, yaGenerado: true, version: COMISIONES_VER,
                 error: 'Las comisiones de ' + mesKey + ' (regla ' + reglaId + ') YA SE GENERARON el ' +
                        (ya[0].timestamp ? _cobStr(_cobD(ya[0].timestamp)) : '—') + ' por ' + (ya[0].usuario || '—') + ': ' +
                        ya.map(function (i) { return i.beneficiario + ' $' + i.monto.toLocaleString('es-MX') + ' (' + i.via + ')'; }).join(', ') +
                        '. No se generó nada de nuevo. Si necesitas rehacerlas, usa «Regenerar» — solo funciona si nadie las ha cobrado todavía.',
                 items: ya };
      }
      var revertidos = [];
      if (ya.length && body.regenerar) {
        // 1) Revisar TODO primero. Si UNA sola pieza ya se cobró, no se toca NADA.
        var estados = ya.map(function (i) { return { item: i, est: _comEstadoRef(i) }; });
        var tocados = estados.filter(function (e) { return e.est.tocado; });
        if (tocados.length)
          return { ok: false, version: COMISIONES_VER, noSePuedeRegenerar: true,
                   error: 'No se puede regenerar ' + mesKey + ': parte de lo generado ya se usó y regenerarlo duplicaría dinero.\n· ' +
                          tocados.map(function (t) { return t.est.motivo; }).join('\n· ') +
                          '\nNo se modificó nada. Corrige a mano lo que sobre, o genera el ajuste como un movimiento aparte.',
                   motivos: tocados.map(function (t) { return t.est.motivo; }) };
        // 2) Nadie lo tocó → revertir.
        estados.forEach(function (e) {
          var i = e.item;
          if (i.via === 'nota_credito') {
            // monto 0 → el crédito sale de "disponible" sin borrar el renglón (deja rastro).
            _cobRegistrarCreditoFavor(i.refId, i.beneficiario, 0, null);
          } else if (i.via === 'efectivo' && e.est.rowNum > 1) {
            var ssD = SpreadsheetApp.openById(EGRESOS_SS_2026);
            var shD = ssD.getSheetByName(EGRESOS_TABS[2026] || 'Egresos2026');
            shD.deleteRow(e.est.rowNum);
            try { CacheService.getScriptCache().remove('gas_egresos_v1_2026'); } catch (_c) {}
          }
          revertidos.push({ beneficiario: i.beneficiario, via: i.via, monto: i.monto, refId: i.refId });
        });
        // 3) Marcar el control (de abajo hacia arriba: borrar filas recorre índices).
        var shC = _comEnsureControl();
        ya.map(function (i) { return i.rowNum; }).sort(function (a, b) { return b - a; })
          .forEach(function (rn) { shC.getRange(rn, 7).setValue('revertida'); });
        try { logAudit(body.usuario || 'sistema', 'Comisiones', 'Revertir', mesKey + '/' + reglaId, '', '', revertidos.length + ' movimiento(s)'); } catch (e) {}
      }

      // ── ESCRIBIR ────────────────────────────────────────────────────────
      // Se escribe UN movimiento POR BENEFICIARIO (calc.totales), no uno por línea de
      // detalle: el coordinador cobra un solo pago al mes por los procedimientos de todos
      // los médicos, no un pago suelto por cada uno.
      var creados = [], errores = [];
      calc.totales.forEach(function (T) {
        var concepto;
        if (calc.modo === 'lista') {
          // Montos exactos de la tabla por tier: no hay % ni base.
          concepto = 'Comisión ' + (calc.regla.nombre || reglaId) + ' ' + mesKey +
                     ' · lista de tarifas (tier ' + (calc.tier != null ? calc.tier : '?') + ') · ' +
                     T.lineas.length + ' procedimiento(s)';
        } else {
          var desglose = T.lineas.map(function (l) { return l.medicoNombre + ' $' + l.base.toLocaleString('es-MX') + ' → $' + l.monto.toLocaleString('es-MX'); }).join('; ');
          var pctTxt = T.lineas.length ? (T.lineas[0].pct + '% × ' + T.lineas[0].parte + '% = ' + T.lineas[0].efectivoPct + '%') : '';
          concepto = 'Comisión ' + (calc.regla.nombre || reglaId) + ' ' + mesKey +
                     ' · escalón ' + pctTxt + ' · ' + desglose;
        }
        if (T.via === 'efectivo') {
          var venc = _comFinDeMes(T.mesAplica);
          var res = saveCxP({ mes: T.mesAplica, proveedor: T.beneficiario,
                              contable: 'Gasto', tipo: 'Variable', subtipo: 'Comisiones',
                              concepto: concepto, monto: T.monto, vencimiento: venc,
                              notas: 'Generada por el módulo de Comisiones (' + reglaId + '/' + mesKey + '). Se paga en efectivo.',
                              usuario: body.usuario || 'sistema' });
          if (!res || !res.ok) { errores.push('CxP de ' + T.beneficiario + ': ' + ((res && res.error) || 'error desconocido')); return; }
          creados.push({ beneficiario: T.beneficiario, via: 'efectivo', monto: T.monto, refId: String(res.id), vencimiento: venc });
        } else if (T.via === 'nota_credito') {
          var op = _comOpCredito(reglaId, mesKey, T.beneficiarioId);
          _cobRegistrarCreditoFavor(op, T.beneficiario, T.monto, _comFinDeMes(mesKey));
          // _cobRegistrarCreditoFavor se traga TODOS sus errores (try/catch vacío) y no
          // devuelve nada: si fallara, escribiríamos en el control un crédito que no existe.
          // Se verifica leyendo de vuelta.
          var chk = _comEstadoRef({ via: 'nota_credito', refId: op, monto: T.monto, beneficiario: T.beneficiario });
          if (chk.tocado || chk.rowNum < 0) { errores.push('Crédito de ' + T.beneficiario + ': no quedó registrado en Creditos_Favor.'); return; }
          creados.push({ beneficiario: T.beneficiario, via: 'nota_credito', monto: T.monto, refId: op, mesAplica: T.mesAplica });
        }
      });

      // Registrar en el control SOLO lo que de verdad se escribió.
      if (creados.length) {
        var shCtl = _comEnsureControl();
        var stamp = new Date();
        shCtl.getRange(shCtl.getLastRow() + 1, 1, creados.length, 9).setValues(creados.map(function (c) {
          return [mesKey, reglaId, c.beneficiario, c.via, c.monto, c.refId, 'activa', body.usuario || 'sistema', stamp];
        }));
      }
      try {
        logAudit(body.usuario || 'sistema', 'Comisiones', body.regenerar ? 'Regenerar' : 'Generar',
                 mesKey + '/' + reglaId, 'Escalón', calc.conteo + ' proc.',
                 calc.pct + '% · ' + creados.length + ' movimiento(s) · $' + calc.totalGeneral);
      } catch (e) {}

      return { ok: errores.length === 0, version: COMISIONES_VER,
               mes: mesKey, reglaId: reglaId, conteo: calc.conteo, pct: calc.pct,
               creados: creados, revertidos: revertidos, errores: errores,
               total: _comRedondea(creados.reduce(function (s, c) { return s + c.monto; }, 0)),
               error: errores.length ? ('Se generaron ' + creados.length + ' movimiento(s), pero fallaron otros:\n· ' + errores.join('\n· ')) : '' };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}

/* ═════════════════════════ SETUP ═════════════════════════ */
function setupComisionesConfig() {
  try {
    var cfg = _comCfg();
    PropertiesService.getScriptProperties().setProperty(COMISIONES_CFG_KEY, JSON.stringify(cfg));
    _comEnsureControl();
    return { ok: true, version: COMISIONES_VER, reglas: (cfg.reglas || []).length, tab: COMISIONES_TAB,
             deploy: _comGuardDeps() || '' };
  } catch (ex) { return { ok: false, error: ex.message, version: COMISIONES_VER }; }
}
