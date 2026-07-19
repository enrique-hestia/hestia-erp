# Resumen General 100% Configurable — Documento de Diseño (VestaOS)

## 1. Visión

El Resumen General deja de ser una pantalla fija y se convierte en un **tablero que cada usuario arma con checkboxes**: elige cualquier dato del ERP, decide si lo ve como KPI, gráfica, tabla o calendario, y lo acomoda arrastrando. Se sentirá profesional porque descansa sobre diez decisiones baratas y consistentes (un solo ritmo de 8px, un solo acento dusty rose, números tabulares, una sola curva de easing, cero salto de layout) más un motor de avisos con animaciones generadas en vivo. Todo reusa la infraestructura que ya existe —charts, drag&drop del editor de menú, recordatorios, poll del chat— sin librerías nuevas y sin tocar la deuda de tokens CSS.

## 2. Experiencia de usuario

### Cómo se ve y se siente

`#view-resumen` se reconstruye como **un contenedor de tablero** con dos capas de estado sobre el MISMO DOM. El widget es la `.card` que ya existe (`:447`), envuelta en una celda de grid. Jerarquía de tres niveles y basta: valor (`--text-xl`/700), título (`--text-base`/600), etiqueta (`--text-xs` muted, uppercase). Money y KPIs siempre con `tabular-nums`. Profundidad discreta: `1px --color-divider` + `--shadow-sm`, hover a `--shadow-md` + `translateY(-2px)`. Una sola curva `var(--transition)` (180ms cubic-bezier(0.16,1,0.3,1)) para todo movimiento.

### Modo Vista vs Modo Edición

Se controla con **una clase en el contenedor** (`.dash-grid.is-edit`), no con re-render pesado.

- **MODO VISTA** (default, 95% del tiempo): limpio, sin cromo. Los widgets solo muestran dato. Al navegar, entran con fade+rise escalonado 40ms cada uno (respetando `prefers-reduced-motion` y `hf_anim_view_resumen` `:17005`). KPIs con count-up rAF ~600ms. Skeletons shimmer reservan altura exacta → CLS 0.
- **MODO EDICIÓN**: entra con el botón "Editar tablero" (visible solo si `hasPermission('config_dashboard')` `:26856`). Aparecen: overlay de malla punteada, un `⋮⋮` de arrastre + toolbar por widget (⚙ configurar · ⤡ tamaño · × quitar), el ghost "＋ Agregar widget", y una **editbar** anclada abajo con estado *dirty* y Guardar/Cancelar. Reusa el patrón `_pcMenuEdit`/`_pcMenuDirty`/`pcMenuSave` (`:26533-26627`), renombrado `_dashEdit`/`_dashDirty`/`dashSave`. `beforeunload` protege cambios sin guardar; Cancelar restaura del snapshot.

### El flujo "arma tu tablero con checkboxes"

Al pulsar "＋ Agregar widget" se abre un panel de dos pasos:

- **Paso 1 — Catálogo de datos**: lista buscable con secciones colapsables por categoría y **un checkbox por dato** (reusa el chip `.cf-chk` `:327`; buscador `.cmdk-inp-wrap` `:364`). Cada fila muestra nombre + micro-badge de unidad + hint de tipo (serie/categoría/escalar). El buscador filtra por label/id/categoría (mismo patrón que el editor de menú `:26546`).
- **Paso 2 — Elegir visual**: según lo marcado se ofrecen **solo las viz compatibles** (un escalar → KPI; una serie → barra o tabla; una distribución → pastel o barra), con preview en vivo. "Agregar al tablero" lo inserta al final del flujo en modo edición.

Estado vacío que enseña: sin layout propio ni global, se muestra un hero "Arma tu tablero" con **tres plantillas de un clic** (Dirección / Finanzas / Operación) usando `.theme-card` (`:264-266`), más un CTA "Empezar de cero". Coach-marks secuenciales la primera vez en Edición (arrastra · redimensiona · agrega), descartables en `localStorage`.

## 3. Arquitectura por subsistema

Idea rectora del tablero: separar **FUENTE** (un fetch), **DATO** (una serie/escalar/distribución extraída) y **WIDGET** (cómo/dónde se dibuja). Un dato nuevo = una entrada de registro; una fuente nueva = un fetcher. Cero reescritura de módulos.

### 3.1 Motor de dashboard

**Cómo funciona.** Tres piezas en el frontend:

- `DASH_SOURCES` — fetchers **deduplicados**: N datos que salen de la misma acción GET comparten UN fetch, cacheado por (fuente + hash de filtros) con TTL. Reusa `getJSON`/`getToken()` (`:26737`); el GET lleva `&token=` explícito porque el wrapper de fetch solo inyecta token en POST (`:26760`).
- `DATA_REGISTRY` — el catálogo declarativo. Cada entrada: `id`, `label`, `categoria`, forma que produce (`serie`/`escalar`/`distribucion`/`tabla`/`calendario`/`avisos`), `render[]` de viz compatibles, `fuente`, `filtros`, y un `get(bundle,filtros)` que extrae la forma normalizada. Añadir un dato = pushear un objeto. Exponer un dato que ya vive en otro dashboard (p.ej. `_aiData.ticketPromedio`) = agregar su fuente + su `get()`, sin tocar `loadAnalisisIngresos()` ni su render.
- Resolver + render: `_dashResolve(widget)` mergea filtros globales↔override, pide el bundle, extrae la forma. `_dashRenderGrid()` mata todos los charts (`_dashKillCharts()`), pinta las tarjetas y resuelve cada widget. **Destruir antes de recrear** es la regla ya establecida (`:2289, :2306, :2342, :2447`); `_dashCharts[wid]` guarda cada instancia. `tabla/kpi/calendario/avisos` son HTML puro (sin canvas → nada que destruir).

Las cuatro formas normalizadas ya existen de facto en `getFallbackData()` (`:2241`): `{meses,ingresos,…}` (serie), `donut:{labels,data,colors}` (distribución), `servicios:[…]` (tabla). El registro solo las estandariza. Los colores salen de `cssVar('--color-primary'|…)` (`:2308-2310`) y las instancias quedan en `Chart.instances`, así `refreshChartColors()` (`:2447`) las recolorea en dark mode sin trabajo extra.

**Grid y reordenamiento.** Modelo A (recomendado para envío): flujo ordenado con spans, estilo Notion. Reordenar = arrastrar con `orden` fraccional (clon literal de `_pcMenuDrag`/`_pcMenuDrop`/`pcMenuSave` `:26593-26627`): soltar sobre otra tarjeta → `D.orden = T.orden + 0.5`, luego renumerar. Redimensionar = ciclar span S/M/L/XL (`grid-column: span 3/4/6/12`) y alto sm/md/lg. Sin colisiones ni matemática de coordenadas. Colapso móvil: todo a `span 12`. Modelo B (grid absoluto 2D estilo Grafana con handles de esquina) queda como Fase 2.

**Dónde vive el dato.** Backend nuevo `dashboard.gs`, hoja `Dashboard_Layouts`. **Una fila = un layout completo** (no una fila por widget): la config de widgets va anidada como blob JSON en la celda `ConfigJSON`. Columna `Usuario` (email lowercase, patrón `recordatorios.gs:15`) o sentinel `__GLOBAL__` para plantillas compartidas (patrón `__AVISO__` de `novedades.gs:31`). Cuál layout está activo para el usuario = preferencia personal → `localStorage` (`hf_dash_active`), no la hoja. `saveDashboardLayout` hace **upsert in-place por (Usuario,LayoutId)** con el patrón `idRow` de `saveMenu` (`core.gs:1208`) — **nunca `sh.clear()`** (la hoja es multi-usuario; sería el bug de blanqueo de `saveSemanalConfig` `semanal.gs:189`). Wiring: `dashboardLayouts` en `doGet` (`core.gs:187`) tras el gate y `_privSet` (`core.gs:279`); `saveDashboardLayout` en `doPost` (`finance.gs:471`) con guard `typeof`. Permiso `config_dashboard` (y `config_dashboard_global` para plantillas), fail-closed.

**Qué se reusa.** `.card`/`.card-header`/`.card-title` (`:447-452`), `.kpi-card` + `drawSpark` (`:426`,`:2287`), `destroy()` antes de `new Chart` + array kill como `_boardKillCharts` (`:21719`), `refreshChartColors()` (`:2447`), el editor de menú completo para DnD/dirty/guardar, `.cf-chk`/`.cmdk-inp-wrap`/`.theme-card` para el constructor.

**Migración sin pérdida.** El Resumen clásico actual (4 KPIs + ingresoChart + donut + tabla servicios + funnel + alertas) se siembra como plantilla `__GLOBAL__` "Resumen clásico" (setup idempotente en `dashboard.gs`), apuntando a datos ya reales. Día 1 el usuario ve lo mismo; luego personaliza. Los bloques que hoy son placeholder de `getFallbackData` (funnel/cashflow/costos/cac) se registran marcados `fuente:'demo'` hasta tener fuente real.

### 3.2 Calendario / pendientes por-usuario

**Cómo funciona.** Widget de primera clase `tipo:'calendario'` (dato `me_pendientes_cal`, categoría Personal). Mini-mes CSS de 7 columnas (sin canvas → nada que destruir): hoy con anillo `box-shadow:0 0 0 2px var(--color-primary)`, días con pendientes con punto `--color-primary`, vencidos (`fecha<hoy && estado==='pendiente'`) en ámbar/rojo. Click en día vacío → `recOpenForm({fecha})` (alta prellenada); click en día con items → panel inline reusando `_recCard(r,hoy)` con ✓Hecho/📅/🗑 ya cableados. Flechas ◀▶ solo re-renderizan `_recData` sin refetch.

**Dónde vive el dato.** Hoja `Recordatorios` existente; GET `action=recordatorios&usuario=&token=` (`recordatorios.gs:142`) ya filtra por email de sesión — cada quien ve solo los suyos. **Cero backend nuevo** (Opción A). Al marcar hecho, `recDone` → `loadRecordatorios` refresca y el motor invalida el bundle `recordatorios|…`.

**Qué se reusa.** `_recData` (`:31721`), `_recCard` (`:31734`), `recOpenForm` (`:31831`), `_recModalHTML` (`:31779`) — hay que incluir `_recModalHTML()` en el DOM del widget para que los ids del form existan.

*Opción B opcional (Fase posterior): `readRecordatoriosMes(usuario,ym)` en `recordatorios.gs` para tilear ocurrencias de recurrencias en el mes usando el motor `_recNextOcc` existente (`:59`); ✓Hecho sobre una ocurrencia virtual sigue llamando `updateRecordatorioEstado(rowNum)` sobre la fila real. Requiere ruta GET `recordatoriosMes` y redeploy.*

> Nota de seguridad heredada (no bloqueante): las acciones de `recordatorios` no pasan por `_tokenHasPermission` y `usuario` es auto-declarado por el frontend. No hay fuga de PII de pacientes (son agendas de staff), pero un usuario podría leer la agenda de otro. Si algún día se cierra, derivar `usuario` de `verifyToken(token)` en el backend.

### 3.3 Avisos (broadcast)

**Cómo funciona.** Los pendientes (por-usuario) y los avisos (broadcast) **nunca se mezclan en almacenamiento**: hojas distintas, endpoints distintos, permisos distintos, colores distintos. El calendario es solo un *renderer* que puede superponer ambas capas. Tres superficies: banner global (franja `--color-primary-hl` con borde-izq 3px, ícono `megaphone`, × descartable recordado en `localStorage`), widget "Avisos del equipo" (`tipo:'avisos'`, categoría Comunicación), y vista de gestión `#view-avisos` para publicar/expirar.

**Dónde vive el dato.** Backend nuevo `avisos.gs` (espejo de `chat.gs`), hoja `Avisos` en `SHEET_ID`. `avisosListActivos(sinceTs, ym, usuario)` es **lectura abierta** (todo el equipo ve broadcasts), filtra `Activo=TRUE`, no expirados, por mes visible y por audiencia. `avisosPublicar`/`avisosActualizar` son **GATED** con `_tokenHasPermission(body.token,'publicar_avisos')` (`finance.gs:5950`), fail-closed, y hacen **update in-place por `Id`** tocando solo columnas presentes (nunca `sh.clear()`, nunca `payload.x||''` — regla dura anti-blanqueo). Wiring: GET `avisosActivos` junto a `core.gs:429`; POST junto a `finance.gs:1137` con guard `typeof`. Un aviso con `FechaEvento` pinta un marcador 🎉 en ese día para TODOS en el calendario.

**Qué se reusa.** El **poll del chat**: extender `chatPoll` (`chat.gs:123`) para adjuntar `res.avisos` y leerlo en `chatPollOnce` (`:19485`) — un solo sondeo cada 7s ya montado, **cero timers nuevos**. El calendario/widget usan el GET dedicado `avisosActivos&ym=` para filtrar por mes. Autor de `getSession()` (`:26716`); token inyectado por el wrapper (`:26760`).

**Permisos.** Un solo permiso nuevo `publicar_avisos`, registrado en tres sitios: `OP_PERMS_CATALOG` (`:26802`, cat "Comunicación"), `ROLE_PERMS_DEFAULT` (`:26841`, fallback admin/director/gerente), y la columna `permisos_operativos` de la hoja `Roles`. El botón "Publicar" solo se pinta con `hasPermission('publicar_avisos')` (`:26856`); el backend igual rechaza el POST a mano.

### 3.4 Animaciones

**Cómo funciona.** No es un mp4 ni video subido: es un **motor de motion-graphics guionado** que renderiza en cliente desde una receta (`AnimTipo` + `AnimParams`), determinista y reproducible ("▶ Ver de nuevo"). Extiende el motor que **ya existe**: el registro `_HF_ANIMS` (`:16786`), `_hfEnsureStyles` idempotente (`:17015`) y **un único bucle RAF disciplinado** `window._hfRAF` con `cancelAnimationFrame` antes de arrancar y auto-stop cuando el nodo se desmonta (`:17029-17046`). La única pieza de sustrato nueva es un renderer canvas-2D para partículas: todo (arte vectorial rasterizado 1 vez, partículas, tipografía con `fillText`) se dibuja en **un solo `<canvas>`**, así el `draw(ctx,params,t)` sirve idéntico para reproducir en vivo y para exportar.

**Dónde vive el dato.** Columna `Anim` de la hoja `Avisos`: blob JSON en una celda (`{tpl, nombre, headline, accent, intensity, motif, seed, duracion}`), patrón key/blob ya usado. `seed` fija el patrón de partículas para que preview == publicado == export. Nada binario se almacena.

**Qué se reusa.** `_HF_ANIMS`/`_hfEnsureStyles`/RAF/`cssVar()` (respeta dark mode automáticamente), el sello `hfSello()` (`:24657`) en cada escena, la fuente Satoshi ya cargada, el poll del chat para transportar el aviso, y el patrón lazy-load de `_loadSheetJS()` (`:4200`) si se activa export GIF.

## 4. Modelo de datos (hojas nuevas)

### Hoja `Dashboard_Layouts` (backend `dashboard.gs`, en su propio libro o SHEET_ID)

| Campo | Tipo | Uso |
|---|---|---|
| `Usuario` | string | email lowercase, o `__GLOBAL__` para plantillas |
| `LayoutId` | string | id del layout (clave junto con Usuario para el upsert) |
| `Nombre` | string | nombre visible ("Mi Resumen") |
| `EsPlantilla` | bool | TRUE = plantilla global compartida |
| `Orden` | number | orden entre layouts del usuario |
| `Activo` | bool | bandera |
| `ConfigJSON` | JSON (celda) | layout completo: grid, filtrosGlobales, array de widgets con su config |
| `ActualizadoEn` | Date | timestamp del upsert |

### Hoja `Avisos` (backend `avisos.gs`, en `SHEET_ID`)

| Campo | Tipo | Uso |
|---|---|---|
| `Id` | `AVISO-00001` | correlativo (patrón `saveRecordatorio:174`) |
| `CreadoEn` | Date | orden y `sinceTs` del poll |
| `AutorEmail` / `AutorNombre` | string | quién publicó (de `getSession()`) |
| `Titulo` / `Mensaje` | string | contenido del broadcast |
| `Nivel` | `info`\|`exito`\|`alerta` | color del banner (brand/verde/ámbar) |
| `FechaEvento` | `YYYY-MM-DD` o vacío | si trae fecha → marcador en el calendario de todos |
| `Activo` | bool | soft-delete (deja registro, no borra) |
| `Expira` | `YYYY-MM-DD` o vacío | auto-oculta pasado ese día |
| `Anim` | vacío o JSON | receta de animación (§5); vacío = sin animación |
| `Segmento` | `todos`\|rol\|`email\|email` | audiencia opcional (default `todos`) |

*El widget-config de cada dashboard NO es una hoja: vive anidado dentro de `ConfigJSON`. Los pendientes NO son hoja nueva: reusan `Recordatorios`.*

## 5. Sistema de animaciones

### Catálogo de plantillas (`_AVISO_TPLS`, cada una un `draw(ctx,p,t)`)

| id | Escena / motion | Ánimo |
|---|---|---|
| `cumple` | pastel dibujándose + velas con flicker, confeti con gravedad+drift, "¡Feliz cumpleaños, {nombre}!" letra por letra | festivo cálido |
| `hito` / `aniversario` | número gigante con reveal por máscara, bokeh dorado, cuenta ascendente | solemne-celebratorio |
| `meta` | 2–3 estallidos de fuegos escalonados + cuenta ascendente del monto (easeOutExpo) + ✓ | triunfal |
| `logro` | medalla/estrella que se ensambla (stroke-draw), rayos radiales, barrido de brillo | prestigio |
| `bienvenida` | telón que abre + nombre entra del centro, chispas bokeh, listón dusty-rose | cálido, elegante |
| `aviso` | **sobrio a propósito**: badge/campana con un pulso, barra ámbar firme, titular fade+rise, **cero confeti** | serio, corporativo (y **fallback**) |
| `generico` | serpentinas suaves de marca + sello VestaOS | neutro-positivo |

Timeline común ~5–6s en fases con stagger (nada de golpe): intro (0–0.6s) → partículas escalonadas (0.6–1.4s) → tipografía con spring/overshoot (1.4–2.2s) → sostener con loop sutil (2.2–4.5s) → outro/reposo. Easing custom (easeOutBack para el nombre, easeOutExpo para números), física real de partículas (gravedad + drift senoidal + rotación), motion blur barato (`rgba` fill en vez de `clearRect`), paleta de marca acotada leída por `cssVar()`, DPR-aware en pantalla y resolución fija en export. `prefers-reduced-motion` → frame clímax estático con fade. Auto-stop del RAF cuando el canvas se desmonta (`if(!document.body.contains(cv)) return;`), pausa si `document.hidden`, degrade adaptativo de conteo de partículas por `hardwareConcurrency`.

### Cómo se auto-eligen desde el texto

Clasificador **sin IA**, 100% cliente y offline, sobre `titulo+mensaje` normalizados (minúsculas, sin acentos). Scoring por prioridad donde **lo sobrio manda**: si matchea `urgente|importante|atencion|mantenimiento|junta|politica|recordatorio` → `aviso` (nunca fiesta sobre algo delicado). Luego categorías celebratorias en orden `cumple > hito > meta > logro > bienvenida`; tono positivo sin categoría → `generico`; si nada → `aviso` low. El resultado **pre-selecciona** el dropdown en el form de publicar (editable), extrae el nombre del homenajeado del directorio de staff y un headline sugerido. Preview en vivo con el mismo canvas antes de publicar.

*Export opcional (Fase posterior): WebM nativo vía `MediaRecorder` + `canvas.captureStream()` (cero dependencias, mismo `draw`); GIF con encoder mini lazy-loaded (~10KB, patrón `_loadSheetJS`); poster PNG con `toDataURL`. MP4 se descarta (exigiría ffmpeg.wasm ~25MB).*

### Nota de privacidad

Todo se genera y reproduce **en el navegador**: cero video subido, cero API externa, cero `webkitSpeechRecognition`, cero autoplay de audio. El clasificador es keyword-matching local. "Cumpleaños" = **staff** por default (lee `_chatDirectorio()` sobre la hoja `Usuarios` `chat.gs:80`, jamás `PACIENTES_SS_ID`). El nombre lo teclea/elige el autor; el JSON `Anim` guarda solo tipo + estilo + ese nombre. Política de contenido: los avisos son broadcast interno y NO deben contener datos de pacientes.

*Ruta Claude API para clasificar queda explícitamente descartada por defecto (respeta el pilar). Si algún día se activara, sería server-side desde Apps Script con key en Script Property, nombre redactado a «NOMBRE», salida JSON de solo parámetros visuales, y opt-in degradable — pero no forma parte de este entregable.*

## 6. Plan por etapas

Cada etapa es un entregable pequeño, byte-idéntico para lo existente, y deja algo útil funcionando.

### E1 — Calendario de pendientes (SOLO-FRONTEND)
Widget/vista mini-calendario reusando `_recData`/`_recCard`/`recOpenForm`/`_recModalHTML` (Opción A, sin backend). Toggle en la vista de Recordatorios y/o como primer widget del Resumen. **Resuelve el requisito de pendientes por-usuario ya**, sin redeploy: bump `<!-- app-build -->`, `node validate-html.js`, commit + `git push`. Es lo pequeño y útil que se puede entregar mañana mismo.

### E2 — Motor de dashboard configurable (FRONTEND + `dashboard.gs` NUEVO / redeploy)
`DASH_SOURCES` + `DATA_REGISTRY` (semilla con datos de `resumenGeneral`/`summary`/`estadoResultados`) + resolver + `_dashRenderGrid`/`_dashKillCharts` + editor DnD + constructor por checkboxes + estado vacío "Arma tu tablero". Backend `dashboard.gs` (`readDashboardLayouts`/`saveDashboardLayout`/`setupDashboard` que siembra "Resumen clásico" `__GLOBAL__`), wiring en `core.gs` doGet + `finance.gs` doPost, permiso `config_dashboard`, y `dashboard.gs` agregado a `$FILES` de `deploy-gas.ps1`. Corre `deploy-gas.ps1` + `setupDashboard()` una vez. **Resuelve los requisitos 1 y 2.**

### E3 — Avisos broadcast (FRONTEND + `avisos.gs` NUEVO / redeploy)
`avisos.gs` (espejo de `chat.gs`: `avisosPublicar`/`avisosActualizar` gated, `avisosListActivos` abierta, `setupAvisos`), permiso `publicar_avisos` en los tres sitios, banner global piggyback en `chatPoll`, form de publicación gated, widget `tipo:'avisos'`, vista `#view-avisos` en la hoja `Menu`. Agregar `avisos.gs` a `$FILES`, `deploy-gas.ps1` + `setupAvisos()` una vez. **Resuelve el requisito de avisos.**

### E4 — Animaciones auto-generadas (SOLO-FRONTEND)
`_AVISO_TPLS` (7 plantillas `draw(ctx,p,t)`) sobre `_HF_ANIMS`/RAF/`cssVar`, clasificador `avisoDetectTpl`, reproductor con overlay + sello `hfSello()`, toggle "Con animación" + preview en el form, botón "▶ Ver" en el banner y en marcadores del calendario. La animación es 100% frontend; el backend de E3 ya persiste el JSON `Anim`. Bump + validate + push. **Resuelve el requisito de "video" auto-generado.**

*(E5 opcional: tileado exacto de recurrencias con `readRecordatoriosMes`, requiere redeploy; Modelo B de grid 2D estilo Grafana; export WebM/GIF/PNG.)*

## 7. Riesgos y decisiones abiertas

**Riesgos técnicos controlados:**
- **Blanqueo de datos** (bug histórico, 4+ incidentes): `saveDashboardLayout` y `avisosActualizar` hacen upsert in-place por clave, **nunca `sh.clear()`** ni `payload.x||''`. Las hojas son multi-usuario/multi-fila; borrar-y-reescribir perdería datos de otros.
- **Fuga de memoria de charts**: `_dashKillCharts()` destruye todas las instancias antes de cada re-render/re-layout; el canvas de animación hace `cancelAnimationFrame` + `canvas.remove()` al cerrar; `tabla/kpi/calendario/avisos` no usan canvas.
- **Deuda de tokens CSS**: usar solo tokens canónicos (`--color-surface-2`, `--color-text`, `--color-divider`, `--color-primary`), nunca `--card-bg`/`--text-primary`/`--text-secondary` (1276 usos con fallback → sin dark mode). Registrar cada Chart para que `refreshChartColors()` lo recoloree.
- **Nada fantasma**: `#view-dashboard`/`#view-avisos` en la hoja `Menu`, gated por permiso; editar layout tras `config_dashboard`, publicar tras `publicar_avisos` (fail-closed en backend con `_tokenHasPermission`).

**Decisiones abiertas:**
1. **Animaciones — profundidad.** El default es motor local por plantilla (offline, sin PII). ¿Se quiere export a video (WebM/GIF) en E4 o queda para después? ¿Se abre alguna vez la ruta Claude API para clasificar (server-side, redactada)? Por defecto **NO**, para no depender de la nube ni rozar el pilar.
2. **Privacidad de cumpleaños.** Confirmar que "cumpleaños" siempre sale del directorio de **staff** (`Usuarios`), jamás de pacientes. ¿Se quiere alguna vez un opt-in de cumpleaños automáticos leyendo un campo `FechaNac` de empleados? Requiere decisión explícita del dueño.
3. **Grid — modelo.** ¿Modelo A (flujo con spans, robusto, reusa DnD) es suficiente, o se quiere el Modelo B (grid 2D libre estilo Grafana) más adelante? El A ya se ve profesional; el B es mucho más código de colisión.
4. **Seguridad de recordatorios.** El `usuario` auto-declarado en las acciones de `recordatorios` permite leer agendas ajenas. No es PII de pacientes, pero conviene decidir si se deriva de `verifyToken(token)` en el backend en algún momento.
5. **Recurrencias en el calendario.** MVP (E1) pinta solo la próxima ocurrencia. ¿Se necesita tileado exacto del mes (E5, requiere `dashboard.gs`/`recordatorios.gs` redeploy) o basta con la próxima?

---

**Archivos a tocar.** Frontend: `hestia-fertility-dashboard.html` (vista `#view-resumen` `:776`, CSS junto a `.card`/`.kpi-grid` `:387-462`, motor + editor + constructor + calendario + avisos + animaciones; bump build + `node validate-html.js` + `git push`). Backend nuevo: `dashboard.gs` y `avisos.gs`, cableados en `core.gs` doGet + `finance.gs` doPost y **agregados a `$FILES` de `deploy-gas.ps1`** (o no se despliegan); correr `deploy-gas.ps1` + `setupDashboard()`/`setupAvisos()` una vez (idempotentes). Permiso `config_dashboard` y `publicar_avisos` en `OP_PERMS_CATALOG` (`:26802`), `ROLE_PERMS_DEFAULT` (`:26841`) y la hoja `Roles`. Items en hoja `Menu` para las vistas nuevas.
