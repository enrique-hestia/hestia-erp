/* ═══════════════════════════════════════════════════════════════════════════
 * TRADUCCIONES — el producto es dueño de su traducción
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * PROBLEMA QUE RESUELVE
 * ---------------------
 * Los reportes que van al paciente (estado de cuenta, cobranza, nota, comprobante
 * de cancelación y carta de seguro) pueden salir en otro idioma, pero los nombres
 * de los productos venían del catálogo SIEMPRE en español. El resultado era un
 * documento mitad inglés / mitad español, que es exactamente de lo que se quejó
 * el cliente. Peor: cuando la traducción fallaba, el documento salía igual, en
 * español y sin avisar.
 *
 * DECISIÓN DE ARQUITECTURA
 * ------------------------
 * La traducción de un producto es un ATRIBUTO DEL PRODUCTO, no del reporte.
 * Vive en BD_Productos, en columnas `Descripcion_EN` / `Descripcion_FR` /
 * `Descripcion_PT` (convención `Descripcion_<LANG>`), agregadas al final de la
 * hoja con `_bdProdColEnsure` (append-safe: nunca reordena ni pisa columnas).
 *
 * El traductor automático (LanguageApp) SOLO SUGIERE. Un humano confirma antes
 * de que nada se guarde. El catálogo es la única fuente de verdad al imprimir:
 * si un producto no tiene traducción confirmada, el documento NO SE IMPRIME.
 *
 * ANTES existían DOS mapas hardcodeados y ya divergentes entre sí
 * (`SEG_TRAD_MAP` en el HTML y `SEG_TRAD_OVERRIDES` en cobranza.gs). Su trabajo
 * curado no se tira: `migrarTraduccionesProductos()` lo vuelca al catálogo una
 * sola vez. A partir de ahí los mapas quedan solo como semilla de arranque —
 * la autoridad es el catálogo.
 * ═══════════════════════════════════════════════════════════════════════════ */

/* Idiomas a los que se puede emitir un documento (además del español, que es el
   original y por definición no necesita traducción). La Carta de Seguro ya
   ofrecía en/fr/pt, así que el sistema NO se hardcodea a un solo idioma. */
var TRAD_LANGS = ['en', 'fr', 'pt'];
var TRAD_LANG_NAMES = { es: 'Español', en: 'English', fr: 'Français', pt: 'Português' };

/* Convención de columna: Descripcion_EN, Descripcion_FR, Descripcion_PT… */
function _tradColName(lang) {
  return 'Descripcion_' + String(lang || '').trim().toUpperCase();
}

function _tradLangOk(lang) {
  return TRAD_LANGS.indexOf(String(lang || '').toLowerCase().substring(0, 2)) > -1;
}

/* Normalización idéntica a la que ya usaba el sistema (_segTrNorm en cobranza.gs
   y _segNorm en el HTML): minúsculas, sin acentos, espacios colapsados. Se
   replica aquí para que traducciones.gs no dependa del orden de carga de
   archivos en Apps Script (scope global compartido, sin imports). */
function _tradNorm(s) {
  s = String(s == null ? '' : s).toLowerCase();
  var map = { 'á':'a','à':'a','ä':'a','é':'e','è':'e','ë':'e','í':'i','ì':'i','ï':'i',
              'ó':'o','ò':'o','ö':'o','ú':'u','ù':'u','ü':'u','ñ':'n' };
  s = s.replace(/[áàäéèëíìïóòöúùüñ]/g, function (c) { return map[c] || c; });
  return s.replace(/\s+/g, ' ').trim();
}

/* ── CONCEPTOS QUE NO SON PRODUCTOS ──────────────────────────────────────────
   No todo lo que sale impreso en un renglón es un producto del catálogo. Una
   cuenta por cobrar con monto suelto sintetiza su línea desde el CONCEPTO, que
   es texto libre (cobranza.gs: `items = [{producto: concepto…}]`, y el concepto
   por default es literalmente 'Saldo inicial'). Ese texto viaja a
   BD_Ingresos.Producto y de ahí al reporte como un concepto más.

   Al imprimir en otro idioma, el embudo (hfTradGate) exige traducción para CADA
   concepto y `saveTraduccionProducto` solo sabía escribir en BD_Productos: un
   concepto de texto libre no empataba con ningún producto → se rechazaba con
   'Producto no encontrado en el catálogo' → el documento NO SALÍA NUNCA, sin
   ninguna salida para el usuario (el mismo botón fallaba una y otra vez).
   Le pasaba hasta al comprobante de una cancelación ya reembolsada.

   Estas traducciones viven en su propia hoja, con la MISMA convención de columna
   (`Descripcion_<LANG>`) y la MISMA normalización (`_tradNorm`) que el catálogo.
   NO se meten a BD_Productos: 'Saldo inicial' no es un producto — no tiene SKU,
   ni precio, ni categoría, ni existencias, y una fila fantasma ahí contaminaría
   los selectores de producto, los reportes de catálogo, el inventario y las
   auditorías. Un no-producto no se arregla inventándole identidad de producto.

   PRECEDENCIA: el CATÁLOGO MANDA. Estas entradas solo RELLENAN HUECOS. Solo se
   escriben cuando no había producto que empatara, así que normalmente no compiten
   con nadie; pero si más adelante alguien da de alta ese texto como producto de
   verdad y le captura su traducción, la del producto debe ganar — si no, editar
   el catálogo no surtiría efecto y el documento imprimiría la traducción vieja
   sin avisar. Rellenar huecos nunca puede empeorar lo que ya estaba bien. */
var TRAD_OVR_TAB = 'Traducciones_Texto';

/* Agrega una columna al final si no existe (match EXACTO de encabezado).
   Mismo patrón append-safe que `_bdProdColEnsure`, pero sobre cualquier hoja. */
function _tradColEnsureOn(sh, headerText) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) { sh.getRange(1, 1).setValue(headerText); return 1; }
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  var idx = hdrs.indexOf(String(headerText).trim().toLowerCase());
  if (idx > -1) return idx + 1;
  sh.getRange(1, lastCol + 1).setValue(headerText);
  return lastCol + 1;
}

/* La hoja de conceptos libres. `crear=false` → null si no existe (lectura: que no
   exista es normal y significa "no hay ninguno", no un error). */
function _tradOvrSheet(crear) {
  var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
  var sh = ss.getSheetByName(TRAD_OVR_TAB);
  if (!sh) {
    if (!crear) return null;
    sh = ss.insertSheet(TRAD_OVR_TAB);
    sh.getRange(1, 1, 1, 3).setValues([['Descripcion', 'Actualizado', 'Usuario']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* { normalizado(Descripcion) → traducción } de los conceptos libres. Nunca tumba
   una impresión: si algo falla, devuelve {} y manda el catálogo solo. */
function _tradOvrMap(lang) {
  try {
    var sh = _tradOvrSheet(false);
    if (!sh) return {};
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {};
    var hdrs = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var iCol = hdrs.indexOf(_tradColName(lang).toLowerCase());
    if (iCol < 0) return {};
    var map = {};
    for (var i = 1; i < data.length; i++) {
      var desc = String(data[i][0] || '').trim();
      var tr   = String(data[i][iCol] || '').trim();
      if (!desc || !tr) continue;
      map[_tradNorm(desc)] = tr;
    }
    return map;
  } catch (e) { return {}; }
}

/* Guarda la traducción de un concepto libre. Empareja por Descripcion
   normalizada (es lo único que trae un reporte) y actualiza en vez de duplicar. */
function _tradOvrSave(lang, desc, traduccion, usuario) {
  var sh  = _tradOvrSheet(true);
  var col = _tradColEnsureOn(sh, _tradColName(lang));
  var data = sh.getDataRange().getValues();
  var normWanted = _tradNorm(desc), target = -1;
  for (var i = 1; i < data.length; i++) {
    if (_tradNorm(data[i][0]) === normWanted) { target = i + 1; break; }
  }
  var anterior = '';
  if (target < 0) {
    target = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(target, 1).setValue(String(desc).trim());
  } else {
    anterior = String(sh.getRange(target, col).getValue() || '');
  }
  sh.getRange(target, col).setValue(traduccion);
  var cAct = _tradColEnsureOn(sh, 'Actualizado');
  var cUsr = _tradColEnsureOn(sh, 'Usuario');
  sh.getRange(target, cAct).setValue(new Date());
  sh.getRange(target, cUsr).setValue(String(usuario || 'sistema'));
  try {
    if (typeof logAudit === 'function') {
      logAudit(usuario || 'sistema', 'Productos', 'Traducir concepto',
               String(desc).trim(), _tradColName(lang), anterior, traduccion);
    }
  } catch (e) { /* la auditoría no debe tumbar el guardado */ }
  return { ok: true, lang: lang, productoId: '', origen: 'concepto',
           descripcion: String(desc).trim(), traduccion: traduccion };
}

/* ── PAÍS → IDIOMA ────────────────────────────────────────────────────────────
   Vive en Config_Dropdowns (sección `Pacientes`, campo `paisIdioma`), NO
   hardcodeado. Ver CFG_DD_DEFAULTS en finance.gs. Formato de los valores:
   `México=es|USA=en`. Se administra desde Panel de Control → Formularios como
   cualquier otro dropdown, y `_ddAddMissingDefaults` lo siembra sin destruir
   nada de lo que ya haya capturado el usuario. */
var TRAD_PAIS_IDIOMA_FALLBACK = ['México=es', 'Mexico=es', 'USA=en', 'Estados Unidos=en'];

function readPaisIdiomaMap() {
  var pares = TRAD_PAIS_IDIOMA_FALLBACK;
  try {
    if (typeof readDropdowns === 'function') {
      var dd = readDropdowns();
      var v = dd && dd.dropdowns && dd.dropdowns['Pacientes'] &&
              dd.dropdowns['Pacientes']['paisIdioma'] &&
              dd.dropdowns['Pacientes']['paisIdioma'].valores;
      if (v && v.length) pares = v;
    }
  } catch (e) { /* se queda con el fallback */ }
  var map = {};
  pares.forEach(function (par) {
    var ix = String(par).indexOf('=');
    if (ix < 1) return;
    var pais = _tradNorm(String(par).substring(0, ix));
    var lang = String(par).substring(ix + 1).trim().toLowerCase().substring(0, 2);
    if (pais && lang) map[pais] = lang;
  });
  return { ok: true, map: map };
}

/* Idioma sugerido por el país. Devuelve '' si el país no está mapeado — el
   llamador decide (normalmente: español, que es el default del sistema). */
function idiomaPorPais(pais) {
  var m = readPaisIdiomaMap().map;
  return m[_tradNorm(pais)] || '';
}

/* ── LECTURA: mapa de traducciones del catálogo ──────────────────────────────
   Devuelve { normalizado(Descripcion) → traducción } para un idioma. Es lo que
   consulta el frontend ANTES de imprimir para decidir si bloquea o no. */
function readTradMap(lang) {
  try {
    lang = String(lang || '').toLowerCase().substring(0, 2);
    if (lang === 'es') return { ok: true, lang: 'es', map: {} }; // el original ya está en español
    if (!_tradLangOk(lang)) return { ok: false, error: 'Idioma no soportado: ' + lang, map: {} };

    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Productos');
    if (!sh) return { ok: false, error: 'BD_Productos no encontrada', map: {} };

    var data = sh.getDataRange().getValues();
    var hdrs = data.length ? data[0].map(function (h) { return String(h).trim().toLowerCase(); }) : [];
    var iCol = hdrs.length ? hdrs.indexOf(_tradColName(lang).toLowerCase()) : -1;

    var map = {};
    if (iCol > -1) {                       // iCol < 0 ⇒ aún nadie ha capturado ninguna
      for (var i = 1; i < data.length; i++) {
        var desc = String(data[i][2] || '').trim();
        var tr   = String(data[i][iCol] || '').trim();
        if (!desc || !tr) continue;
        map[_tradNorm(desc)] = tr;
      }
    }
    // Conceptos que no son productos (ver TRAD_OVR_TAB). RELLENAN HUECOS: nunca
    // pisan al catálogo, que es y sigue siendo la autoridad.
    var ovr = _tradOvrMap(lang), nOvr = 0;
    for (var k in ovr) { if (ovr.hasOwnProperty(k) && !map[k]) { map[k] = ovr[k]; nOvr++; } }

    return { ok: true, lang: lang, map: map, sinColumna: iCol < 0, conceptos: nOvr };
  } catch (ex) {
    return { ok: false, error: ex.message, map: {} };
  }
}

/* ── ESCRITURA: guardar la traducción de un producto ─────────────────────────
   Se llama desde el alta/edición del producto en el Catálogo y también desde la
   pantalla de bloqueo (el usuario tiene un cliente enfrente: bloquear no puede
   significar abandonarlo, así que puede capturarla ahí mismo y desbloquear).

   Empareja por ProductoID si viene; si no, por Descripcion normalizada — que es
   lo único que trae un reporte. */
function saveTraduccionProducto(body) {
  try {
    body = body || {};
    var lang = String(body.lang || '').toLowerCase().substring(0, 2);
    if (!_tradLangOk(lang)) return { ok: false, error: 'Idioma no soportado: ' + lang };
    var traduccion = String(body.traduccion || '').trim();
    if (!traduccion) return { ok: false, error: 'La traducción no puede ir vacía.' };

    var prodId = String(body.productoId || '').trim();
    var desc   = String(body.descripcion || '').trim();
    if (!prodId && !desc) return { ok: false, error: 'Falta productoId o descripcion.' };

    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Productos');
    if (!sh) return { ok: false, error: 'BD_Productos no encontrada' };

    var col = _bdProdColEnsure(sh, _tradColName(lang)); // append-safe
    var data = sh.getDataRange().getValues();
    var target = -1, normWanted = _tradNorm(desc);
    for (var i = 1; i < data.length; i++) {
      if (prodId) { if (String(data[i][0]).trim() === prodId) { target = i + 1; break; } }
      else if (normWanted && _tradNorm(data[i][2]) === normWanted) { target = i + 1; break; }
    }
    // ── SIN MATCH EN EL CATÁLOGO ─────────────────────────────────────────────
    // No siempre es un error: un reporte también imprime CONCEPTOS LIBRES que
    // nunca fueron productos ('Saldo inicial', el concepto de una cuenta por
    // cobrar con monto suelto…). Devolver error aquí dejaba al usuario ATRAPADO:
    // el embudo exige la traducción para imprimir, y guardarla era imposible →
    // el documento no salía nunca por más veces que le diera al botón.
    // Se guarda en la hoja de conceptos, que readTradMap fusiona al imprimir.
    // Si vino productoId, sí es un error real: se pidió un producto CONCRETO por
    // su id y ese id no existe — inventarle un concepto libre taparía el bug.
    if (target < 0) {
      if (prodId) return { ok: false, error: 'Producto no encontrado en el catálogo: ' + prodId };
      if (!desc)  return { ok: false, error: 'Falta la descripción del concepto a traducir.' };
      return _tradOvrSave(lang, desc, traduccion, body.usuario);
    }

    var anterior = String(sh.getRange(target, col).getValue() || '');
    sh.getRange(target, col).setValue(traduccion);
    try {
      if (typeof logAudit === 'function') {
        logAudit(body.usuario || 'sistema', 'Productos', 'Traducir',
                 String(data[target - 1][0] || ''), _tradColName(lang), anterior, traduccion);
      }
    } catch (e) { /* la auditoría no debe tumbar el guardado */ }
    return { ok: true, lang: lang, origen: 'catalogo', productoId: String(data[target - 1][0] || ''),
             descripcion: String(data[target - 1][2] || ''), traduccion: traduccion };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}

/* Escribe las traducciones que vengan en el body de un alta/edición de producto
   (`body.traducciones = {EN:'…', FR:'…'}`). Solo toca los idiomas presentes:
   omitir uno NO lo borra. Se llama desde saveNewProducto/createProducto/
   updateProducto (finance.gs) para que el formulario del Catálogo sea el lugar
   natural donde nace la traducción. */
function _bdProdWriteTrads(sh, rowNum, body) {
  var trads = body && body.traducciones;
  if (typeof trads === 'string') { try { trads = JSON.parse(trads); } catch (e) { trads = null; } }
  if (!trads || typeof trads !== 'object') return 0;
  var n = 0;
  TRAD_LANGS.forEach(function (l) {
    var val = trads[l.toUpperCase()];
    if (val === undefined) val = trads[l];
    if (val === undefined) return;          // idioma no enviado → no se toca
    var col = _bdProdColEnsure(sh, _tradColName(l));
    sh.getRange(rowNum, col).setValue(String(val == null ? '' : val).trim());
    n++;
  });
  return n;
}

/* Guardado en lote — la pantalla de bloqueo suele traer varios faltantes. */
function saveTraduccionesProductoBatch(body) {
  try {
    body = body || {};
    var items = body.items;
    if (typeof items === 'string') { try { items = JSON.parse(items); } catch (e) { items = []; } }
    if (!items || !items.length) return { ok: false, error: 'Nada que guardar.' };
    var guardados = 0, errores = [], conceptos = 0;
    for (var i = 0; i < items.length; i++) {
      var r = saveTraduccionProducto({
        lang: items[i].lang || body.lang,
        productoId: items[i].productoId,
        descripcion: items[i].descripcion,
        traduccion: items[i].traduccion,
        usuario: body.usuario
      });
      if (r.ok) { guardados++; if (r.origen === 'concepto') conceptos++; }
      else errores.push((items[i].descripcion || items[i].productoId || '(sin descripción)') + ': ' + (r.error || 'error desconocido'));
    }
    // Un fallo PARCIAL debe ser legible. Antes esto devolvía `ok:false` a secas y el
    // front (`if(!d||!d.ok) throw`) lo mostraba como un error mudo que además borraba
    // el hecho de que N traducciones SÍ se guardaron — el usuario reintentaba todo
    // sin saber qué había quedado. Ahora `error` viene siempre armado y dice ambas
    // cosas: cuántas entraron y exactamente cuál falló y por qué.
    var ok = errores.length === 0;
    return { ok: ok, guardados: guardados, conceptos: conceptos, errores: errores,
             error: ok ? '' : ('Se guardaron ' + guardados + ' de ' + items.length + '. No se pudo guardar: ' + errores.join(' · ')) };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}

/* ── SIEMBRA: no tirar el trabajo curado ─────────────────────────────────────
   Los dos mapas hardcodeados traían ~36 términos médicos ya revisados por un
   humano. Esto los vuelca al catálogo UNA vez.

   `TRAD_SEED_FRONTEND` es la copia del `SEG_TRAD_MAP` que vivía en el HTML (el
   backend no puede leer el frontend). Se fusiona con `SEG_TRAD_OVERRIDES`
   (cobranza.gs, mismo scope global), que gana en los empates por estar mejor
   curado (usa nombres completos tipo "In Vitro Fertilization (IVF)").

   IDEMPOTENTE: si la celda ya tiene algo capturado, NO se pisa. Correrla dos
   veces no duplica ni sobreescribe nada.

   MATCH EXACTO sobre la descripción normalizada, a propósito: el viejo
   `_segTradLocal` hacía match por substring ("consulta" dentro de cualquier
   descripción que la contuviera), que para SUGERIR al vuelo es tolerable pero
   para SEMBRAR un catálogo que un humano va a dar por bueno es peligroso —
   metería traducciones incorrectas con cara de confirmadas. Lo que no empata
   exacto se reporta como pendiente, no se adivina. */
var TRAD_SEED_FRONTEND = {
  en: {
    'fertilizacion in vitro':'In Vitro Fertilization','fiv':'In Vitro Fertilization (IVF)','fiv icsi':'IVF with ICSI','icsi':'ICSI',
    'histeroscopia':'Hysteroscopy','histeroscopia diagnostica':'Diagnostic Hysteroscopy','histeroscopia quirurgica':'Surgical Hysteroscopy',
    'criopreservacion':'Cryopreservation','anualidad criopreservacion':'Cryopreservation Annuity','criopreservacion de embriones':'Embryo Cryopreservation','criopreservacion de ovulos':'Egg Cryopreservation',
    'estimulacion ovarica controlada':'Controlled Ovarian Stimulation','estimulacion ovarica':'Ovarian Stimulation',
    'transferencia de embriones':'Embryo Transfer','transferencia de embrion':'Embryo Transfer','transferencia embrionaria':'Embryo Transfer',
    'consulta':'Consultation','consulta de valoracion':'Assessment Consultation','consulta subsecuente':'Follow-up Consultation','valoracion':'Assessment',
    'almacenamiento':'Storage','ovodonacion':'Egg Donation','donacion de ovulos':'Egg Donation',
    'inseminacion artificial':'Artificial Insemination','inseminacion intrauterina':'Intrauterine Insemination','induccion de ovulacion':'Ovulation Induction',
    'biopsia embrionaria':'Embryo Biopsy','medicamentos':'Medications','medicamento':'Medication','laboratorio':'Laboratory',
    'estudios':'Tests','estudio':'Test','ultrasonido':'Ultrasound','perfil hormonal':'Hormone Panel','espermatobioscopia':'Semen Analysis','seguimiento folicular':'Follicular Monitoring'
  },
  fr: {
    'fertilizacion in vitro':'Fécondation in vitro','fiv':'Fécondation in vitro (FIV)','icsi':'ICSI','histeroscopia':'Hystéroscopie',
    'criopreservacion':'Cryoconservation','anualidad criopreservacion':'Annuité de cryoconservation','estimulacion ovarica controlada':'Stimulation ovarienne contrôlée',
    'transferencia de embriones':'Transfert d’embryons','consulta':'Consultation','almacenamiento':'Stockage','ovodonacion':'Don d’ovocytes','medicamentos':'Médicaments','laboratorio':'Laboratoire'
  },
  pt: {
    'fertilizacion in vitro':'Fertilização in vitro','fiv':'Fertilização in vitro (FIV)','icsi':'ICSI','histeroscopia':'Histeroscopia',
    'criopreservacion':'Criopreservação','anualidad criopreservacion':'Anuidade de criopreservação','estimulacion ovarica controlada':'Estimulação ovariana controlada',
    'transferencia de embriones':'Transferência de embriões','consulta':'Consulta','almacenamiento':'Armazenamento','ovodonacion':'Doação de óvulos','medicamentos':'Medicamentos','laboratorio':'Laboratório'
  }
};

/* Semilla fusionada para un idioma: frontend como base, backend (mejor curado) encima. */
function _tradSeedFor(lang) {
  var out = {};
  var base = TRAD_SEED_FRONTEND[lang] || {};
  for (var k in base) { if (base.hasOwnProperty(k)) out[_tradNorm(k)] = base[k]; }
  var over = {};
  try { over = (typeof SEG_TRAD_OVERRIDES !== 'undefined' && SEG_TRAD_OVERRIDES[lang]) || {}; } catch (e) { over = {}; }
  for (var k2 in over) { if (over.hasOwnProperty(k2)) out[_tradNorm(k2)] = over[k2]; }
  return out;
}

function migrarTraduccionesProductos(body) {
  try {
    body = body || {};
    var langs = body.lang ? [String(body.lang).toLowerCase().substring(0, 2)] : TRAD_LANGS.slice();
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Productos');
    if (!sh) return { ok: false, error: 'BD_Productos no encontrada' };

    var resumen = {};
    langs.forEach(function (lang) {
      if (!_tradLangOk(lang)) { resumen[lang] = { error: 'Idioma no soportado' }; return; }
      var seed = _tradSeedFor(lang);
      var col  = _bdProdColEnsure(sh, _tradColName(lang)); // append-safe
      var data = sh.getDataRange().getValues();

      var escritos = 0, yaTenian = 0, sinSemilla = [], vals = [];
      for (var i = 1; i < data.length; i++) {
        var actual = String(data[i][col - 1] || '').trim();
        var desc   = String(data[i][2] || '').trim();
        var id     = String(data[i][0] || '').trim();
        if (actual) { yaTenian++; vals.push([actual]); continue; }   // IDEMPOTENTE: no pisa
        if (!id || !desc) { vals.push(['']); continue; }
        var hit = seed[_tradNorm(desc)];                              // MATCH EXACTO
        if (hit) { vals.push([hit]); escritos++; }
        else { vals.push(['']); sinSemilla.push(desc); }
      }
      // Una sola escritura (evita cientos de llamadas a Sheets y el límite de 6 min)
      if (vals.length) sh.getRange(2, col, vals.length, 1).setValues(vals);
      resumen[lang] = {
        columna: _tradColName(lang),
        sembrados: escritos,
        yaTenian: yaTenian,
        sinTraduccion: sinSemilla.length,
        ejemplosSinTraduccion: sinSemilla.slice(0, 15)
      };
    });
    return { ok: true, resumen: resumen };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}

/* ── AUDITORÍA (SOLO LECTURA) ────────────────────────────────────────────────
   Qué productos activos NO tienen traducción todavía. Para que el usuario lo
   sepa ANTES de toparse con el bloqueo con un cliente enfrente. NO escribe. */
function auditarTraduccionesProductos(body) {
  try {
    body = body || {};
    var langs = body.lang ? [String(body.lang).toLowerCase().substring(0, 2)] : TRAD_LANGS.slice();
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Productos');
    if (!sh) return { ok: false, error: 'BD_Productos no encontrada' };
    var data = sh.getDataRange().getValues();
    if (!data.length) return { ok: true, resumen: {} };
    var hdrs = data[0].map(function (h) { return String(h).trim().toLowerCase(); });

    var resumen = {};
    langs.forEach(function (lang) {
      if (!_tradLangOk(lang)) { resumen[lang] = { error: 'Idioma no soportado' }; return; }
      var iCol = hdrs.indexOf(_tradColName(lang).toLowerCase());
      var faltan = [], conTrad = 0, activos = 0;
      for (var i = 1; i < data.length; i++) {
        var id = String(data[i][0] || '').trim();
        if (!id) continue;
        var activo = !(data[i][6] === false || String(data[i][6]).toUpperCase() === 'FALSE');
        if (!activo) continue;
        activos++;
        var tr = iCol > -1 ? String(data[i][iCol] || '').trim() : '';
        if (tr) { conTrad++; continue; }
        faltan.push({ productoId: id, sku: String(data[i][1] || ''),
                      descripcion: String(data[i][2] || ''), categoria: String(data[i][3] || '') });
      }
      resumen[lang] = {
        columna: _tradColName(lang), existeColumna: iCol > -1,
        activos: activos, conTraduccion: conTrad, faltan: faltan.length, detalle: faltan
      };
    });
    return { ok: true, resumen: resumen };
  } catch (ex) {
    return { ok: false, error: ex.message };
  }
}

/* ── SETUP ───────────────────────────────────────────────────────────────────
   1) Crea las columnas Descripcion_<LANG> en BD_Productos (append-safe).
   2) Agrega la columna `Idioma` a la hoja de Pacientes (append-safe): así
      `readPacienteFull` la devuelve sola (lee la fila entera por encabezado) y
      `insertRow`/`updateRow` la escriben solas (mapean por encabezado).
   3) Siembra los dropdowns nuevos (Pacientes|idioma, Pacientes|paisIdioma) vía
      el `setupConfigDropdowns()` que ya existe — no destructivo. */
function setupTraducciones() {
  var pasos = [];
  try {
    var ss = SpreadsheetApp.openById(PRODUCTOS_SS_ID);
    var sh = ss.getSheetByName('BD_Productos');
    if (!sh) return { ok: false, error: 'BD_Productos no encontrada' };
    TRAD_LANGS.forEach(function (l) {
      var c = _bdProdColEnsure(sh, _tradColName(l));
      pasos.push('BD_Productos.' + _tradColName(l) + ' → columna ' + c);
    });
  } catch (e) { return { ok: false, error: 'BD_Productos: ' + e.message, pasos: pasos }; }

  try {
    var col = _pacColEnsure('Idioma');
    pasos.push('Pacientes.Idioma → columna ' + col);
  } catch (e2) { pasos.push('⚠ Pacientes.Idioma: ' + e2.message); }

  try {
    if (typeof setupConfigDropdowns === 'function') {
      var r = setupConfigDropdowns();
      pasos.push('Config_Dropdowns: ' + (r && r.msg ? r.msg : 'ok'));
    }
  } catch (e3) { pasos.push('⚠ Config_Dropdowns: ' + e3.message); }

  return { ok: true, pasos: pasos,
           msg: 'Listo. Ahora corre migrarTraduccionesProductos() para sembrar los términos ya curados.' };
}

/* Agrega una columna a la hoja de Pacientes si no existe (match EXACTO de
   encabezado, se añade al final). Mismo patrón append-safe que
   `_bdProdColEnsure`, pero sobre el spreadsheet de Pacientes. */
function _pacColEnsure(headerText) {
  var ss = SpreadsheetApp.openById(PACIENTES_SS_ID);
  var sh = ss.getSheetByName('Pacientes') || ss.getSheets()[0];
  var lastCol = sh.getLastColumn();
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var want = String(headerText).trim().toLowerCase();
  var idx = hdrs.indexOf(want);
  if (idx > -1) return idx + 1;
  sh.getRange(1, lastCol + 1).setValue(headerText);
  return lastCol + 1;
}
