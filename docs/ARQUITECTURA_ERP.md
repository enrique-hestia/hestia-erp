<!-- Generado 2026-07-18 · mapeo builds .238-.246 · plano para la migración, NO producción -->

# Arquitectura VestaOS ERP

> Documento maestro de arquitectura de datos para migrar VestaOS (clínica de fertilidad Hestia) desde Google Sheets + Apps Script hacia una base de datos relacional destino. Escrito para que un ingeniero humano construya la BD y migre **sin romper nada**, en paralelo, módulo por módulo.
>
> **Versión del mapeo:** builds .238–.240. **Alcance:** dominios finanzas, bancos, catálogos y nómina.

---

## 0. Cómo leer este documento

- Cada tabla lleva: columnas → tipo Postgres, PK/FK, índices, si es **CALIENTE** (alto volumen) y qué columnas son **hoy fórmula de Sheets / recálculo en código** (marcadas 🔥-TRAMPA).
- Convención de trampa: cualquier columna marcada **[DERIVADA]** hoy se calcula al vuelo o con fórmula de celda. En Postgres se resuelve como **columna generada**, **vista**, **vista materializada** o **trigger** — nunca como número suelto capturado, salvo que se indique "materializar por rendimiento".
- Todo dato de paciente es **[SENSIBLE]**: la privacidad de pacientes es un pilar y el enmascarado deja de ser cosmético (de pantalla) para vivir en la BD (RLS + vistas).

---

## 1. Resumen y decisiones

### 1.1 Stack elegido

| Capa | Decisión | Por qué |
|---|---|---|
| **Base de datos** | **PostgreSQL 15+** (gestionado con **Supabase**) | Relacional real, transacciones ACID, `jsonb` nativo para los campos JSON embebidos, columnas generadas, particionamiento declarativo por rango (año) y por lista (tenant), window functions para saldos corridos, y **Row-Level Security** de primera clase — el núcleo de la seguridad de pacientes. |
| **Auth** | **Supabase Auth** (GoTrue) | `auth.uid()` disponible dentro de las políticas RLS; soporta el portal de pacientes (acceso externo) y el staff con el mismo motor. Contraseñas hasheadas (bcrypt/argon2) por el proveedor — nunca en tabla propia en texto. |
| **API** | **PostgREST** (incluido en Supabase) + funciones RPC (`plpgsql`) para operaciones compuestas (cancelaciones, conciliación, comisiones) | Reemplaza el `doGet`/`doPost` de Apps Script. Las RLS aplican igual vía API que vía SQL directo, así que el enmascarado no se puede saltar desde el cliente. |
| **Frontend** | HTML/JS actual servido en **Cloudflare Pages** | Resuelve el hosting pendiente (el ISP del usuario bloquea github.io); dominio propio `app.hestiafertility.com`. Estático, sin build. El fetch cambia de `APPS_SCRIPT_URL` a la URL de Supabase con `anon key`. |
| **Almacenamiento de archivos** | **Supabase Storage** (o dejar los hipervínculos a Drive durante la transición) | Comprobantes, CFDIs, cotizaciones. Los `*_url` de hoy siguen apuntando a Drive hasta migrar los binarios. |
| **Jobs** | Supabase **cron (pg_cron)** / Edge Functions | Recálculo de comisiones MP mensuales, generación de suscripciones crío, refresh de vistas materializadas. Sustituye los triggers de tiempo de Apps Script. |

### 1.2 Principios de diseño (criterios del dueño, citados)

1. **"ERP de verdad, escalable a >1,000,000 de movimientos/año sin romperse ni hacerse lento."** → Tablas calientes particionadas por rango de fecha (año) + índices compuestos por `(tenant, fecha)` y `(tenant, entidad)`. Nada de "devolver TODO": toda lectura de tabla caliente es paginada por keyset. Saldos y reportes materializados, no recalculados en cada request.
2. **"Estabilidad garantizada; datos de pacientes con seguridad máxima."** → RLS obligatoria en toda tabla con PII; enmascarado movido a la BD (vistas + políticas por rol); PII sensible (CURP, NSS, CLABE, RFC) cifrable a nivel columna; auditoría append-only.
3. **"Vendible/replicable a cualquier giro comercial (multi-tenant)."** → Columna `tenant_id` en todas las tablas + RLS por tenant, con opción de escalar a **una base por cliente** para aislamiento fuerte. Núcleo reusable (finanzas/bancos/CxP/CxC/inventario/nómina MX) separado de lo específico de fertilidad (ciclos/crío/lab).
4. **"Habrá portal de pacientes (acceso externo)."** → El paciente autenticado ve **solo sus propias filas** (`paciente_id = auth.uid()` vía tabla puente), nunca las de otros ni datos internos. Se diseña desde el día uno, no se parcha después.

### 1.3 Decisión de fondo: una tabla con `tenant_id` + partición, no un libro por año

El modelo actual de **un spreadsheet por año y por módulo** (`INGRESOS_SS_2024/25/26`, `Egresos2024/25/26`) es un particionamiento manual frágil: si el año no está configurado, el código **cae en silencio al libro 2026** y lee/escribe el año equivocado (`_ingIdDeAnio`/`_egIdDeAnio`). En Postgres esto desaparece: **una sola tabla con columna `anio` (o mejor, particionada por rango de `fecha`)**. El año deja de ser un artefacto de archivo y pasa a ser un dato.

---

## 2. Esquema de datos

### 2.0 Convenciones globales

Toda tabla incorpora, además de las columnas de negocio:

```sql
tenant_id     uuid        NOT NULL REFERENCES tenants(id),   -- multi-tenant
id            bigint      GENERATED ALWAYS AS IDENTITY,       -- surrogate estable
created_at    timestamptz NOT NULL DEFAULT now(),
updated_at    timestamptz NOT NULL DEFAULT now(),
created_by    text,                                          -- email/usuario
deleted_at    timestamptz                                    -- soft-delete donde aplique
```

- **Nunca se hace hard-delete** en tablas de dinero: los soft-delete de hoy (`abonos_cxp.reversado`, `comisiones_generadas.estado='revertida'`, ventas `cancelada`) se conservan como `deleted_at`/`estado`.
- Las columnas de folio de negocio (OP-00000, CXP-00000, PROD-00001) se **conservan como columna de texto** para continuidad visual, pero la identidad interna es siempre `id bigint`. Los folios se generan con secuencias, **no** escaneando el máximo de la columna (que hoy se hace por bug de duplicados en Sheets).

---

### DOMINIO: FINANZAS

#### 2.1 `bd_ingresos` — 🔥 LA TABLA MÁS CALIENTE DEL ERP · CALIENTE

Ledger de ventas. Una OP abarca varias líneas. Hoy vive en un libro por año.

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | surrogate (PK real interna) |
| `op` | text | folio OP-00000. **No único por sí solo** |
| `linea` | int | línea dentro de la OP |
| `anio` | int | clave de partición (deriva de `fecha`) |
| `fecha` | date | de aquí se derivan mes/mesIdx en código (no se guardan) |
| `paciente_id` | bigint FK → pacientes | **[SENSIBLE]** — enmascarado a "OP" sin permiso `ver_datos_sensibles` |
| `paciente_nombre` | text | denormalizado histórico; **[SENSIBLE]** |
| `categoria` | text | |
| `producto` | text | **FK lógica** a bd_productos por nombre/descr (no por id) |
| `pvp` | numeric(14,2) | precio de lista |
| `descuento` | numeric(14,2) | |
| `cantidad` | numeric(14,2) | |
| `total_pagar` | numeric(14,2) | base de casi todos los KPIs; **excluir filas canceladas** |
| `pagado` | numeric(14,2) **NULLABLE** | 🔥-TRAMPA: **vacío ⇒ pagado por completo** (histórico). Ver §2.1.1 |
| `monto_fact_mes` | numeric(14,2) | facturación del mes |
| `forma_pago` | text | rutea a banco/comisión vía `config_formas_pago` |
| `facturacion` | boolean | |
| `conciliacion` | boolean | |
| `contabilidad` | boolean | |
| `observaciones` | text | |
| `factura` | text | folio/etiqueta factura |
| `poliza` | text | póliza contable (ContaDigital) |
| `usmx` | text | moneda USD/MXN |
| `ciclo_alta_baja` | text | |
| `sucursal` | text | |
| `archivo_url` | text | hipervínculo a comprobante en Drive |
| `razon_social` | text | **[SENSIBLE]** receptor fiscal del CFDI |
| `origen_externo` | text | **FK lógica** → origenes_externos (motor de comisiones) |
| `factura_rfc` | text | **[SENSIBLE]** |
| `factura_uuid` | text | UUID CFDI. **FK lógica** → cfdi emitido |
| `pagos_detalle` | **jsonb** | abonos `[{...}]` |
| `cancelada` | boolean | venta cancelada se conserva, **no suma en KPIs** |
| `cancelacion_data` | **jsonb** | traza del reverso |

- **PK lógica de negocio:** `UNIQUE (tenant_id, anio, op, linea)`.
- **FKs:** `producto → bd_productos` (por nombre; ver §2.1.2), `origen_externo → origenes_externos`, `factura_uuid → cfdi_emitidos`. Referenciada por `cuentas_cobrar.op`, `abonos_cobrar.op`, `creditos_favor.op`, `creditos_consumo.op`.
- **Particionamiento:** `PARTITION BY RANGE (fecha)` con partición por año. Reemplaza el "un libro por año" de facto, sin el riesgo de leer el año equivocado.
- **Índices:**
  ```sql
  CREATE INDEX ON bd_ingresos (tenant_id, op);
  CREATE INDEX ON bd_ingresos (tenant_id, fecha);
  CREATE INDEX ON bd_ingresos (tenant_id, paciente_id);
  CREATE INDEX ON bd_ingresos (tenant_id, origen_externo) WHERE origen_externo IS NOT NULL;
  CREATE INDEX ON bd_ingresos (tenant_id, factura_uuid) WHERE factura_uuid IS NOT NULL;
  CREATE INDEX ON bd_ingresos (tenant_id) WHERE cancelada IS NOT TRUE; -- KPIs excluyen canceladas
  ```

##### 2.1.1 🔥-TRAMPA `pagado` semi-nulo con semántica

`pagado` VACÍO ⇒ la línea se asume **pagada por completo** (dato histórico). Un `num()` ingenuo la leería como 0 y **fabricaría deuda falsa**. Regla en Postgres:

```sql
-- saldo real de la línea
saldo := total_pagar - COALESCE(pagado, total_pagar);
```

Mantener la columna **nullable** y usar `COALESCE(pagado, total_pagar)` en todo cálculo. **No** rellenar los nulos con `total_pagar` en la migración: se perdería la señal "vacío = histórico pagado". La cobranza (Motor A) **no** deriva saldos de aquí — usa el registro explícito de `cuentas_cobrar` (§2.6).

##### 2.1.2 FK por nombre (deuda técnica a resolver)

`producto` liga a `bd_productos` **por texto**, no por `ProductoID`. Al migrar: conservar el texto pero **añadir `producto_id bigint` resuelto** por match durante la migración, con reporte de los no-emparejados. No romper el texto (hay históricos con nombres que ya no existen en catálogo).

---

#### 2.2 `egresos` — ledger de gastos **Y** cuentas por pagar (misma fila) · CALIENTE

🔥-TRAMPA #4: **CxP y egreso son la misma fila.** `fecha` VACÍA = CxP (por pagar); `fecha` con valor = egreso pagado. Migrar como **una sola tabla** con estatus derivado, no dos.

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | (hoy `id` es consecutivo **por año**; global aquí) |
| `folio_anio` | int | año del folio original (para trazabilidad) |
| `anio` | int | partición |
| `fecha` | date **NULLABLE** | 🔥 vacía = CxP; con fecha = pagado |
| `mes` | text | 'Mmm-YY' derivado — **[DERIVADA]**, no fecha real |
| `prioridad` | int | col D real de CxP |
| `proveedor` | text/FK | **[SENSIBLE, enmascarable a socio]** → proveedores |
| `contable` | text | clasificación (Gasto/Costo…) |
| `tipo` | text | Fijo/Variable |
| `subtipo` | text | |
| `concepto` | text | **[SENSIBLE, enmascarable]** texto libre |
| `egresos` | numeric(14,2) | el monto del gasto/CxP |
| `notas` | text | enmascarable |
| `vencimiento` | date | |
| `facturacion` | boolean | |
| `pagado` | boolean | TRUE cuando la CxP se saldó (`pagarCxP` o saldo de abonos = 0) |
| `contabilidad` | boolean | |
| `poliza` | text | ContaDigital |
| `forma_pago` | text | rutea a banco (`config_formas_pago`) |
| `observaciones` | text | |
| `link_factura` | text | hipervínculo Drive |
| `link_pago` | text | hipervínculo Drive (col T/20) |
| `cotizacion` | text | hipervínculo Drive |
| `monto_usd` | numeric(14,2) | |
| `tipo_cambio` | numeric(14,6) | |
| `divisa` | text | MXN/USD |
| `estatus` | text | 'Cancelada' para orden cancelada |
| `recurrente_id` | text/FK | → gastos_fijos |
| `mes_devengado` | text | yyyy-MM (informativo, no toca flujo) |
| `mp_comision_id` | text | 🔥 `MPCOM-YYYY-MM`: partida **automática** de comisiones MP |
| `factura_uuid` | text | CFDI del proveedor |
| `factura_folio` | text | SERIE-FOLIO |
| `factura_rfc` | text | **[SENSIBLE, enmascarable]** |
| `factura_razon_social` | text | **[SENSIBLE, enmascarable]** |

- **Estatus derivado (columna generada):**
  ```sql
  estatus_flujo text GENERATED ALWAYS AS (
    CASE WHEN estatus = 'Cancelada' THEN 'cancelada'
         WHEN fecha IS NULL         THEN 'por_pagar'   -- CxP
         ELSE 'pagado' END) STORED;
  ```
- **FKs:** `proveedor → proveedores`, `recurrente_id → gastos_fijos`. Referenciada por `abonos_cxp.cxp_id`, `creditos_proveedor.cxp_id_origen`, `comisiones_generadas.ref_id` (vía efectivo), `ordenes_compra.CxPId`.
- **Índices:** `(tenant_id, proveedor)`, `(tenant_id, fecha)`, `(tenant_id, pagado)`, `(tenant_id, estatus_flujo)`, `(tenant_id, mp_comision_id) WHERE mp_comision_id IS NOT NULL`, partición por año.

##### 🔥-TRAMPA saldo de CxP

No hay columna de saldo. Saldo de una CxP = `egresos − Σ(abonos_cxp activos NO reversados)`. Se resuelve con **vista** `v_cxp_saldos` (ver §7). `abonos_cxp.reversado` es soft-delete: nunca borra.

---

#### 2.3 `bd_cxp` — esquema paralelo secundario

Definido (`setupBDCxP`) pero en la práctica las CxP viven en la fila de `egresos` (fecha vacía). **Recomendación:** **no migrar como tabla propia**; fusionar en `egresos`. Se documenta por completitud (PK `CXP-00000`, FK `proveedor → proveedores`, columnas espejo de egresos + `fecha_registro timestamptz`). Si el ingeniero decide conservarla, tratarla como vista sobre `egresos WHERE estatus_flujo='por_pagar'`.

---

#### 2.4 `creditos_proveedor` — saldo a favor tras cancelar orden pagada

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `fecha` | date | |
| `proveedor` | text/FK | **[SENSIBLE]** — un crédito SOLO aplica al mismo proveedor (blindaje) |
| `monto` | numeric(14,2) | monto original del crédito |
| `monto_disponible` | numeric(14,2) | 🔥-TRAMPA: **saldo vivo** (running balance) mantenido en código |
| `origen` | text | por qué nació (orden cancelada #…) |
| `cxp_id_origen` | bigint FK → egresos.id | |
| `usuario` | text | |
| `notas` | text | |

🔥-TRAMPA #5: `monto_disponible` **no es fórmula**: se decrementa al consumir y se restaura al reversar. Regenerarlo a ciegas lo **regresa a su monto original → dinero duplicado**. En Postgres:

```sql
-- disponible = monto − Σ consumos, NUNCA un número suelto
monto_disponible numeric(14,2) GENERATED ALWAYS AS ... -- si es viable como generated
-- o vista:
SELECT c.id, c.monto - COALESCE(SUM(a.monto),0) AS disponible
FROM creditos_proveedor c
LEFT JOIN abonos_cxp a ON a.credito_id = c.id AND a.reversado IS NOT TRUE
GROUP BY c.id, c.monto;
```

Referenciada por `abonos_cxp.credito_id`.

---

#### 2.5 `abonos_cxp` — ledger de pagos parciales de CxP · MEDIO

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `fecha` | date | |
| `cxp_id` | bigint FK → egresos.id | orden que se abona |
| `proveedor` | text | **[SENSIBLE]** |
| `concepto` | text | |
| `monto` | numeric(14,2) | |
| `origen` | text | `'pago'` \| `'credito'` |
| `credito_id` | bigint FK → creditos_proveedor.id | cuando origen='credito' |
| `forma_pago` | text | rutea a banco/caja |
| `usuario` | text | |
| `notas` | text | |
| `reversado` | boolean | 🔥 **soft-delete**. Saldo = monto − Σ(abonos activos) |

- **Índices:** `(tenant_id, cxp_id)`, `(tenant_id, credito_id)`.

---

#### 2.6 `cuentas_cobrar` — Motor A de cobranza (saldos EXPLÍCITOS) · MEDIO

Registro explícito de adeudos; **no** se infiere de `bd_ingresos.pagado`.

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | (hoja sin PK propia) |
| `fecha` | date | |
| `op` | text FK → bd_ingresos.op | puede ir vacío en saldos iniciales tipo agencia |
| `paciente_id` | bigint FK → pacientes | **[SENSIBLE]** |
| `categoria` | text | |
| `concepto` | text | |
| `monto_cargo` | numeric(14,2) | 🔥-TRAMPA #6: si `nota='auto-ingreso'` es **[DERIVADA]** (Facturado − Pagado de la OP), reescrita por trigger de captura/edición del ingreso. **No editable a mano** |
| `estatus` | text | Pendiente/Pagado/Cancelado (cerrado ≠ borrado) |
| `nota` | text | `'auto-ingreso'` = fila gobernada por la venta (idempotente por OP) |
| `usuario` | text | |
| `timestamp` | timestamptz | |
| `items` | **jsonb** | partidas `[{producto,cantidad,precio,total,pac,fecha}]` |

- **Unicidad lógica:** `UNIQUE (tenant_id, op) WHERE nota='auto-ingreso'` (idempotente por OP).
- 🔥 Para las filas `auto-ingreso`, `monto_cargo` se modela como **trigger** `AFTER INSERT/UPDATE ON bd_ingresos` que recalcula `Facturado − COALESCE(pagado,total)` por OP. Saldo real por renglón = `monto_cargo − Σ abonos_cobrar de la OP`.

---

#### 2.7 `abonos_cobrar` — pagos recibidos contra saldo/suscripción · MEDIO

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `fecha` | date | |
| `op` | text FK → bd_ingresos.op | |
| `paciente_id` | bigint FK → pacientes | **[SENSIBLE]** |
| `monto` | numeric(14,2) | |
| `forma_pago` | text | |
| `banco` | text | banco destino del abono |
| `tipo` | text | 🔥 `'abono-op'` \| `'abono-cargo'` \| `'suscripcion'`. abono-op/abono-cargo **ya** dejaron el saldo canónico en `cuentas_cobrar` → **NO** restar de nuevo por OP (evita doble conteo) |
| `nota` | text | |
| `usuario` | text | |
| `timestamp` | timestamptz | |

- **Índices:** `(tenant_id, op)`, `(tenant_id, paciente_id)`.

---

#### 2.8 `suscripciones_crio` — Motor B cobranza (vencimientos criopreservación)

Se **genera** desde Inventario Crío + config, no se captura fila por fila.

| Columna | Tipo | Notas |
|---|---|---|
| `suscripcion_id` | text | PK de negocio |
| `id` | bigint identity | surrogate |
| `fecha` | date | vencimiento del periodo |
| `paciente_id` | bigint FK → pacientes | **[SENSIBLE]**. Universo = pacientes con inventario crío |
| `plan` | text | anual/mensual |
| `periodo_inicio` | date | |
| `periodo_fin` | date | |
| `periodo` | text | etiqueta ('SALDO-INICIAL' consolida backlog en 1 renglón) |
| `monto` | numeric(14,2) | 🔥-[DERIVADA] de config (tarifaAnual 5700 / tarifaMensual 475) + reglas (año 1 gratis solo Hestia) |
| `estatus` | text | Pendiente/… |
| `abonado` | numeric(14,2) | acumulado pagado; saldo = monto − abonado |
| `concepto` | text | |
| `usuario` | text | |
| `timestamp` | timestamptz | |

- La **config de tarifas** hoy vive en Script Property `COBRANZA_CONFIG` → tabla `config_cobranza` o key-value (§2.24).

---

#### 2.9 `suscripciones_mp` — ligas paciente↔suscripción recurrente MP

| Columna | Tipo | Notas |
|---|---|---|
| `paciente_key` | text | PK lógica normalizada |
| `id` | bigint identity | |
| `paciente_id` | bigint FK → pacientes | **[SENSIBLE]** |
| `payer_id` | text | id del pagador MP (idempotencia) |
| `mp_email` | text | **[SENSIBLE]** |
| `mp_status` | text | |
| `mp_plan` | text | |
| `mp_monto` | numeric(14,2) | |
| `mp_inicio` / `mp_proxcobro` / `mp_ultcobro` | date | |
| `mp_ultmonto` | numeric(14,2) | |
| `mp_cobros` | int | conteo |
| `clasificacion` | text | |
| `actualizado_en` | timestamptz | |
| `actualizado_por` | text | |

- Todas las columnas son append-safe (`_mpColEnsure`) → en Postgres columnas normales, migrar **por nombre**.
- **Unicidad:** `UNIQUE (tenant_id, payer_id)`.

---

#### 2.10 `creditos_favor` — crédito a favor del cliente/médico (upsert por OP)

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `fecha` | date | |
| `op` | text | 🔥 **UPSERT por OP** (OP sintética `COM-regla-mes-medico` para comisiones). FK → bd_ingresos.op |
| `paciente_id`/`beneficiario` | text | **[SENSIBLE]** (o médico beneficiario) |
| `monto_credito` | numeric(14,2) | 🔥-TRAMPA #5: **saldo vivo absoluto** (upsert). Consumir decrementa; 0 lo cierra sin borrar. No es fórmula |
| `nota` | text | `'credito-favor'` / `'credito-favor-nc-reverso'` |
| `usuario` | text | |
| `timestamp` | timestamptz | |

- **Unicidad:** `UNIQUE (tenant_id, op)`. Referenciada por `comisiones_generadas.ref_id` (vía nota_credito).

---

#### 2.11 `creditos_consumo` — consumo de notas de crédito

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `fecha` | date | |
| `op` | text FK → bd_ingresos.op | |
| `paciente_id` | bigint | **[SENSIBLE]** |
| `monto_nc` | numeric(14,2) | monto de la NC |
| `aplicado` | numeric(14,2) | cuánto se aplicó |
| `excedente` | numeric(14,2) | sobrante que regresa a crédito a favor |
| `autorizado` | text | |
| `usuario` | text | |
| `timestamp` | timestamptz | |

---

#### 2.12 `descuentos_agencia` — escala de descuento por volumen (config)

| Columna | Tipo | Notas |
|---|---|---|
| `agencia` | text FK → origenes_externos | |
| `producto` | text | |
| `desde` | int | escalón cantidad desde |
| `hasta` | int | '' = sin tope → `NULL` (interpretar como ∞) |
| `descuento_pct` | numeric(6,3) | |

- **PK:** `(tenant_id, agencia, producto, desde)`. Config → migrando a JSON en Script Property `COBRANZA_DESC_CFG`; en Postgres puede quedar como tabla de config (preferible a JSON para consultarla).

---

#### 2.13 `comisiones_generadas` — idempotencia de comisiones por volumen

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `mes` | text | yyyy-MM |
| `regla_id` | text | FK → config comisiones (Script Property) |
| `beneficiario` | text | médico/coordinador — **[SENSIBLE: nombre]** |
| `via` | text | `'nota_credito'` \| `'efectivo'` |
| `monto` | numeric(14,2) | 🔥-[DERIVADA] (base_medico × pct_escalón × parte%) |
| `ref_id` | text | FK: OP sintética en creditos_favor (nota_credito) o egresos.id (efectivo) |
| `estado` | text | `'activa'` \| `'revertida'` (soft-delete; base de idempotencia) |
| `usuario` | text | |
| `timestamp` | timestamptz | |

- **Unicidad de las activas:** `UNIQUE (tenant_id, mes, regla_id, beneficiario, via) WHERE estado='activa'`.
- La **config de comisiones** NO es tabla: Script Property `COMISIONES_CFG` (JSON de reglas/escalones/beneficiarios) → `config_comisiones` (§2.24).

---

#### 2.14 Bancos (finanzas) → ver DOMINIO BANCOS (§2.15–2.19), modelo unificado

Los `banco_santander` / `banco_amex` / `banco_mercadopago` / `caja_chica` del dominio finanzas son **la misma realidad** que el dominio bancos describe con más detalle. **Se implementa el modelo unificado del dominio bancos** (`config_cuentas` + `movimiento_bancario` + `caja_chica_movimiento`). No duplicar tablas por banco.

---

### DOMINIO: BANCOS (modelo unificado, autoritativo)

#### 2.15 `config_cuentas` — catálogo de cuentas bancarias (config, frío)

| Columna | Tipo | Notas |
|---|---|---|
| `key` | text | **PK**. slug normalizado: 'santander','amex','mercadopago'. Los 3 core nunca deben faltar |
| `nombre` | text | visible |
| `gid` | int | legado Sheets (pestaña). **Deprecar tras migración** |
| `tipo` | text CHECK | discriminador ÚNICO: `'bancaria'`\|`'credito'`\|`'tpv'`. Gobierna layout + comportamiento del saldo |
| `color` | text | hex UI |
| `activo` | boolean | default true |

- **Volumen:** ~3–10 filas de por vida. Sin índices extra.

#### 2.16 `movimiento_bancario` — unifica 3 layouts en 1 tabla · CALIENTE

Fusiona Santander (bancaria) / AMEX (credito) / Mercado Pago (tpv) con columnas específicas de tipo **nullable**.

| Columna | Tipo | Aplica a tipo | Notas |
|---|---|---|---|
| `id` | uuid/bigint | todos | PK sintética (hoy identidad = banco+fila, frágil) |
| `cuenta_key` | text FK → config_cuentas.key | todos | reemplaza el ruteo por gid |
| `seq` | bigint | todos | 🔥 **CRÍTICO**: orden de inserción dentro de la cuenta. El saldo corrido corre por **orden de fila (append), no por fecha** |
| `fecha` | date NULLABLE | todos | vacía en placeholders |
| `mes` | text | tpv/credito | 'Jul-2026' texto histórico, **no date**. Agrupa comisiones MP |
| `deposito` | numeric(14,2) | bancaria | sube saldo |
| `retiro` | numeric(14,2) | bancaria | baja saldo |
| `movimiento` | numeric(14,2) | credito | monto con signo; suma directa (deuda) |
| `cobro` | numeric(14,2) | tpv | bruto cobrado |
| `comisiones` | numeric(14,2) | tpv | comisión MP; base de la partida `MPCOM-YYYY-MM` |
| `pct_comision` | numeric(8,6) | tpv | 🔥-[DERIVADA] = 1−(total_venta/cobro), **fracción** no % entero |
| `total_venta` | numeric(14,2) | tpv | NETO. Base del saldo corrido MP y de liberado/por-liberar |
| `saldo_corrido` | numeric(14,2) | todos | 🔥-TRAMPA PRINCIPAL (§2.16.1) |
| `liberado` | boolean | tpv | TRUE = cuenta al saldo; FALSE con total>0 = 'por liberar' |
| `referencia` | text | todos | **[SENSIBLE, enmascarable]** (`_privRef`) |
| `usd` | numeric(14,2) | bancaria/credito | |
| `tipo_cambio` | numeric(12,6) | bancaria/credito | |
| `poliza` | text | todos | ContaDigital |
| `observaciones` | text | todos | **[SENSIBLE, enmascarable]** |
| `tipo_mov` | text | todos | bancaria='deposito'/'retiro', credito='cargo'/'pago', tpv='CARGO'/'PAGO'/'REVERSO' |

- **CHECK por tipo** que exija las columnas correctas y prohíba las ajenas:
  ```sql
  CONSTRAINT chk_bancaria CHECK (tipo <> 'bancaria' OR (movimiento IS NULL AND cobro IS NULL)),
  CONSTRAINT chk_credito  CHECK (tipo <> 'credito'  OR (deposito IS NULL AND retiro IS NULL AND cobro IS NULL)),
  CONSTRAINT chk_tpv      CHECK (tipo <> 'tpv'      OR (deposito IS NULL AND retiro IS NULL AND movimiento IS NULL))
  ```
  (el `tipo` se resuelve por join a config_cuentas o se denormaliza).
- **Índices:** `(tenant_id, cuenta_key, seq)`, `(tenant_id, cuenta_key, fecha)`, `(tenant_id, mes)` para recálculo de comisiones MP. **El saldo corrido es el punto de dolor de rendimiento.**

##### 2.16.1 🔥-TRAMPA saldo corrido (la trampa central de TODO el sistema)

`saldo_corrido` **no es dato capturado**: hoy se recomputa por tipo y por **orden de fila**:
- bancaria: `Σ(deposito − retiro)`
- credito: `Σ(movimiento)`
- tpv: `Σ(total_venta)`

En Postgres, **dos opciones** (elegir por rendimiento):
1. **Vista con window** (correcto, simple):
   ```sql
   SUM(delta) OVER (PARTITION BY cuenta_key ORDER BY seq
                    ROWS UNBOUNDED PRECEDING)
   -- delta = COALESCE(deposito,0)-COALESCE(retiro,0)  (bancaria)
   --       = COALESCE(movimiento,0)                    (credito)
   --       = COALESCE(total_venta,0)                   (tpv)
   ```
2. **Columna materializada mantenida por trigger** (si la cuenta tiene cientos de miles de filas y se consulta el saldo actual muy seguido).

**Reglas duras:**
- **Persistir `seq`** (orden de captura). Ordenar por `fecha` da saldos distintos y equivocados.
- El **saldo disponible MP** solo suma filas `liberado=TRUE`; 'por liberar' = `total_venta>0 AND liberado=FALSE`.
- **Descartar filas placeholder** en la migración (MP y Caja Chica dejan filas reservadas con solo la fórmula de saldo y fecha/concepto vacíos): no son movimientos reales.

#### 2.17 `caja_chica_movimiento` · MEDIO

| Columna | Tipo | Notas |
|---|---|---|
| `id` | uuid/bigint | PK |
| `anio` | int | pestaña por año → columna |
| `seq` | bigint | orden de fila (el TOTAL corre por captura, no por fecha) |
| `fecha` | date NULLABLE | 🔥 mezcla DD/MM/YYYY y MM/DD/YYYY: desambiguar por >12 al cargar (`parseFechaMx`) |
| `concepto` | text | |
| `es_remanente` | boolean | TRUE si concepto empieza con 'REMANENTE' (siembra el saldo inicial) |
| `salida` | numeric(14,2) | egreso de caja |
| `entrada` | numeric(14,2) | ingreso/reposición |
| `total` | numeric(14,2) | 🔥-[DERIVADA]: window SUM(entrada−salida) por seq arrancando del REMANENTE |

#### 2.18 `config_formas_pago` — ruteo forma→banco+comisión (config, frío)

Vive en el libro **principal** (SHEET_ID), no en el de bancos — al unificar, tabla normal.

| Columna | Tipo | Notas |
|---|---|---|
| `forma_pago` | text | parte de PK. 'Efectivo','Santander','Transferencia','AMEX','TDC','TDD','Mercado Pago' (normalizada) |
| `contexto` | text CHECK | parte de PK. `'ingreso'`\|`'egreso'`\|`'ambos'`. 🔥 **la misma forma rutea distinto según contexto** (AMEX ingreso→MP+comisión / egreso→amex) |
| `banco_key` | text | FK suave → config_cuentas.key **+ valor especial `'cajachica'`** |
| `aplica_comision` | boolean | |
| `comision_pct` | numeric(8,6) NULLABLE | 🔥 vacío ('') = usa default del banco; `''` ≠ 0 |
| `activo` | boolean | default true |

- **PK:** `(tenant_id, forma_pago, contexto)`. ~14 filas semilla.
- **FK caja chica:** 'cajachica' no está en config_cuentas → o se agrega como cuenta especial tipo `'caja'`, o el CHECK de `banco_key` admite el literal `'cajachica'`.

---

### DOMINIO: CATÁLOGOS

#### 2.19 `bd_productos` — catálogo ÚNICO (venta + inventario + fiscal)

Solo las 8 primeras columnas son de posición fija; el resto son **dinámicas por encabezado** (crear como columnas normales, migrar **por nombre** no por orden).

| Columna | Tipo | Notas |
|---|---|---|
| `ProductoID` | text | **PK** de negocio. PROD-00001 |
| `id` | bigint identity | surrogate |
| `SKU` | text | **UNIQUE**. Prefijo por categoría (MED-0001). FK desde Movimientos_Inventario/Combos/OC_Lineas |
| `Descripcion` | text | nombre comercial |
| `Categoria` | text | 'Medicamento' → inventariable automático |
| `Tipo` | text | subcategoría de medicamento |
| `Notas` | text | REPROVIDA/GRUPOMED marca lista especial |
| `Activo` | boolean | readProductos filtra activos |
| `FechaCreacion` | timestamptz | |
| `Inventariable` | boolean | dinámica; default false |
| `Unidad` | text | dinámica (match exacto de encabezado) |
| `StockMinimo` / `StockMaximo` | numeric(14,2) | dinámica |
| `CostoUnitario` | numeric(14,2) | dinámica; valorInventario = StockActual×CostoUnitario |
| `ProveedorPreferido` | text | dinámica |
| `StockActual` | numeric(14,2) | 🔥-TRAMPA: **saldo corrido escrito SOLO por el ledger** (§2.19.1) |
| `ClaveProdServ`, `ClaveUnidad`, `UnidadTexto`, `ObjetoImp`, `TipoFactor`, `TasaIVA` | text | fiscales CFDI 4.0 dinámicas |
| `Descripcion_EN/FR/PT` | text | traducciones dinámicas (autoridad única del nombre extranjero) |

- **Volumen:** catálogo estable (~cientos), frío.

##### 2.19.1 🔥-TRAMPA StockActual (invariante de inventario)

`StockActual` **no es fórmula** y **nunca** se edita desde formulario: es un saldo materializado que **solo** escribe el motor `_registrarMovimientoInventario`, en la **misma transacción** que agrega la fila a `movimientos_inventario` (con lock). La fuente de verdad es el **ledger**. En Postgres: **trigger** `AFTER INSERT ON movimientos_inventario` que actualiza `StockActual`, o vista de saldo. Si no se respeta esta invariante, el stock se corrompe.

#### 2.20 `bd_precios` — historial de precios por lista

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `ProductoID` | text FK → bd_productos | |
| `VigenciaDesde` | date | precio vigente = la más reciente ≤ hoy; futuras se ignoran |
| `Precio` | numeric(14,2) | limpiar $ y , al leer |
| `Moneda` | text | MXN default; USD Surrogacy |
| `ModificadoPor` | text | |
| `FechaModificacion` | timestamptz | |
| `Lista` | text FK → bd_listas | 🔥 **col 7 SIN encabezado** en la hoja. Vacío = 'General'. Legacy 'GM' → 'GrupoMedico' |

- **Unicidad natural:** `(tenant_id, ProductoID, Lista, VigenciaDesde)`. `updateProducto` hoy borra+reescribe una fila por producto+lista → en la práctica 1 fila viva por par, pero **conservar el histórico 2025/2026** de la migración.

#### 2.21 `bd_listas` — listas de precio (config)

| Columna | Tipo | Notas |
|---|---|---|
| `Lista` | text | **PK**. Defaults: General, GrupoMedico, Surrogacy, REPROVIDA. 'General' no borrable |
| `Descripcion` | text | |
| `Moneda` | text | MXN/USD |
| `Multiplicador` | numeric(10,4) | factor sobre base (default 1) |
| `Activo` | boolean | solo activas alimentan dropdown de Pacientes |

#### 2.22 `bd_auditoria` — log append-only · crece indefinido

| Columna | Tipo | Notas |
|---|---|---|
| `id` | bigint identity | |
| `Timestamp` | timestamptz | |
| `Usuario` | text | 'sistema' si no viene |
| `Modulo` / `Accion` / `Referencia` / `Campo` / `Anterior` / `Nuevo` | text | |

- **Estrategia:** partición por `Timestamp` (mensual) + política de retención/archivado. Ver §5.4. Este patrón se generaliza a una **tabla de auditoría global** (§3.5).

#### 2.23 `pacientes` — tabla SENSIBLE (pilar de privacidad)

| Columna | Tipo | Notas |
|---|---|---|
| `ID` | text | **PK** de negocio (col A) |
| `id` | bigint identity | surrogate |
| `nombre_completo` | text | **[SENSIBLE]** (col B fija). Antiduplicado por nombre |
| `fecha_nacimiento` | date | **[SENSIBLE]** |
| `email` | text | **[SENSIBLE, enmascarable]** (col D fija). Antiduplicado por correo (dice qué paciente lo tiene) |
| `origen`, `canal`, `medico_tratante`, `pais`, `idioma` | text | **[SENSIBLE parcial]**; precargan atribución/idioma de comprobantes |
| `lista_precios` | text FK → bd_listas | col J fija (PAC_COL_LISTA) |
| `observaciones` | text | **[SENSIBLE]** |
| `razon_social`, `rfc`, `codigo_postal`, `uso_cfdi`, `regimen_fiscal`, `forma_pago_habitual` | text | fiscales **[SENSIBLE, enmascarable]** |

- **Antiduplicado:** `UNIQUE (tenant_id, lower(nombre_completo))` y `UNIQUE (tenant_id, lower(email)) WHERE email <> ''` (informar cuál colisiona, no solo rechazar).
- **Portal de pacientes:** tabla puente `paciente_auth (paciente_id, auth_user_id)` para ligar el `auth.uid()` de Supabase a su fila. Ver §3.3.
- **Índices:** `(tenant_id, lower(nombre_completo))`, `(tenant_id, lower(email))`.

#### 2.24 `movimientos_inventario` — ledger de stock · CALIENTE

| Columna | Tipo | Notas |
|---|---|---|
| `ID` | text | **PK**. MOV-yyyyMMddHHmmss-nnn |
| `id` | bigint identity | |
| `Fecha` | date | |
| `Modulo` | text | guarda la Categoría (Medicamento/Insumos) para filtrar sin catálogos separados |
| `SKU` | text FK → bd_productos.SKU | |
| `Nombre` | text | descripción cacheada |
| `Tipo` | text | Entrada/Salida/Ajuste (Ajuste = saldo absoluto deseado) |
| `Cantidad` | numeric(14,2) | abs para E/S; delta para Ajuste |
| `Motivo` | text | Compra/Venta/Sobrante/Merma/Caducado/Ajuste Manual |
| `Referencia` | text | OP de venta, ID de OC, etc. |
| `CostoUnitario` | numeric(14,2) | |
| `SaldoResultante` | numeric(14,2) | 🔥 saldo corrido escrito por el sistema (refleja StockActual) |
| `Usuario` / `Notas` | text | |

- **Índices:** `(tenant_id, SKU, Fecha)`. **El lector actual topa en 500 filas** — en Postgres, paginación keyset. Alto volumen.

#### 2.25 `combos` — BOM (config)

| Columna | Tipo | Notas |
|---|---|---|
| `ID` | text | **PK**. COMBO-… |
| `ProductoIngresos` | text | producto de venta que dispara descuento — **match por NOMBRE** (case-insensitive), no SKU/ID |
| `SKU` | text FK → bd_productos.SKU | componente a descontar |
| `NombreMedicamento` | text | cacheado |
| `CantidadPorUnidad` | numeric(14,4) | |
| `Activo` | boolean | borrado lógico = FALSE |

#### 2.26 `ordenes_compra` · MEDIO

| Columna | Tipo | Notas |
|---|---|---|
| `ID` | text | **PK**. OC-00001 |
| `Fecha` | date | |
| `Proveedor` | text | |
| `EstadoRecepcion` | text | 'Pendiente de recibir'/'Recibida' (recepción y pago son flags independientes) |
| `FechaRecepcion` | date | |
| `Total` | numeric(14,2) | suma de líneas |
| `CxPId` | bigint FK → egresos.id | 🔥 toda OC genera su CxP vía saveCxP |
| `CxPRowNum` | int | **[DEPRECAR]**: hoy apunta a la fila de Egresos2026; el estado Pagado se lee en vivo de ahí. En Postgres se resuelve por FK `CxPId` + join |
| `Usuario` / `Notas` | text | |

- 🔥 El estado **Pagado NO vive aquí**: se resuelve en vivo contra `egresos.pagado` por la FK. Una sola fuente de verdad.

#### 2.27 `ordenes_compra_lineas`

| Columna | Tipo | Notas |
|---|---|---|
| `OrdenID` | text FK → ordenes_compra.ID | |
| `SKU` | text FK → bd_productos.SKU | |
| `Nombre` | text | cacheado |
| `Cantidad` | numeric(14,2) | |
| `CostoUnitario` | numeric(14,2) | |
| `Subtotal` | numeric(14,2) | Cantidad×CostoUnitario ([DERIVADA]) |

- **Natural:** `(OrdenID, SKU)`.

#### 2.28 Hojas de LAB y MEDICAMENTOS (captura clínica — específico de fertilidad)

Estas son **captura/referencia clínica**, no catálogo administrable. `lab.gs` lee por **posición fija** hardcodeada (mover una columna rompe el lector). Al migrar: normalizar a columnas nombradas. **Todas son específicas del giro fertilidad** (ver §4).

- **`lab_art`** (ciclos ART): `mes_anio`, `fecha`, `oocytes` int, `2PN` (col aún no configurada, `-1`), `blastocistos` (idem). Paciente por nombre no normalizado. Encabezados fusionados fila 1+2, datos desde fila 3.
- **`lab_fet`** (transferencias): `fecha`, `survived` ('Si'/'No'), `beta`, `clinical_preg`. Fila 1 vacía.
- **`lab_inventario_crio`** (banco criogénico, **[SENSIBLE]**): `paciente_nombre` (col A), `fecha_crio`, `oov` int, `emb` int. **Universo de suscripciones crío.**
- **`lab_insumos`**: `fecha` (col C), `insumo`, `proveedor`, `costo`. Captura suelta, sin ledger.
- **`med_estimulacion`**: `fecha`, `paciente` **[SENSIBLE]**, `cancelado` boolean (🔥 **marcador estructural**: todas las columnas a la derecha de 'Cancelado' son medicamentos-cantidad, ancho variable), `costo_meds`, `<medicamentos...>`. En Postgres: modelar los fármacos como **filas** en una tabla hija `med_estimulacion_detalle(estimulacion_id, medicamento, cantidad)`, no como columnas que crecen a la derecha.
- **`med_ent_med`** (legacy compras de medicamento): `fecha`, `medicamento`, `cantidad`, `precio_unitario`, `total`. Se está sustituyendo por `ordenes_compra`.

---

### DOMINIO: NÓMINA (libro PROPIO — NOMINA_SS_ID, no SHEET_ID)

> ⚠ La nómina vive en un **libro propio** (Script Property `NOMINA_SS_ID`). La hoja "Empleados_RESPALDO_NO_USAR" de SHEET_ID es un respaldo huérfano — **no migrar esa**. El mapeo entregado se corta en `nomina_bonos`; las tablas completas conocidas son las cuatro núcleo + bonos + SBC (esta última descrita a partir del contexto de memoria).

#### 2.29 `nomina_empleados` — la MÁS sensible por PII · frío

| Columna | Tipo | Notas |
|---|---|---|
| `NumEmpleado` | text | **PK** de negocio |
| `Nombre` | text | **[PII]** |
| `RFC` | text | **[PII, cifrable]** |
| `CURP` | text | **[PII MUY sensible, cifrable]** |
| `NSS` | text | **[PII MUY sensible, cifrable]** IMSS |
| `Puesto`, `Departamento` | text | (Departamento = Área; lo llena también el CFDI) |
| `FechaIngreso` | date | antigüedad/vacaciones LFT 2023 |
| `FechaNacimiento` | date | **[PII]** (append-safe) |
| `SalarioDiario` | numeric(14,2) | **[SENSIBLE]**. Siembra SueldoBase = SalarioDiario × días |
| `SBC` | numeric(14,2) | **[SENSIBLE]** salario base de cotización IMSS |
| `Periodicidad` | text | semanal\|quincenal\|mensual (mixto); vacío = default config |
| `Tipo` | text | Sueldos vs asimilado |
| `Banco` | text | **[SENSIBLE]** |
| `CLABE` | text | **[PII MUY sensible, cifrable]** |
| `UsuarioEmail` | text FK → usuarios(email) | habilita "Mis recibos". **[PII]** |
| `Activo` | boolean | 'Sí'/'No' → boolean |
| `Notas` | text | |

- **Cifrado a nivel columna** (pgcrypto / Supabase Vault) para CURP, NSS, CLABE, RFC. Ver §3.4.

#### 2.30 `nomina_periodos` — calendario · frío (~89 filas/año)

| Columna | Tipo | Notas |
|---|---|---|
| `PeriodoID` | text | **PK**. MEN-2026-01 \| QNA-2026-01-1 \| SEM-2026-01 \| EXT-… |
| `Tipo` | text CHECK | semanal\|quincenal\|mensual (enum) |
| `FechaInicio` / `FechaFin` / `FechaPago` | date | FechaPago default = FechaFin |
| `Estatus` | text CHECK | borrador\|validada\|pagada ('pagada' congela captura a solo lectura) |
| `Notas` | text | resumen de órdenes CxP generadas |
| `CreadoEn` | timestamptz | |
| `CreadoPor` | text | |
| `TipoNomina` | text CHECK | ORDINARIA\|EXTRAORDINARIA (append-safe) |

#### 2.31 `nomina_captura` — captura/conciliación · CALIENTE en captura

PK compuesta `(PeriodoID, EmpleadoID)`, upsert por empleado×periodo. FKs a periodos y empleados.

Columnas capturadas: `SueldoBase` **[SENSIBLE]**, `Bonos`, `ValesDespensa`, `Combustible`, `PrimaVacacional`, `OtrasPercepciones`, `OtrasDeducciones`.

Columnas **[DERIVADAS]** (hoy fórmula de Sheets → columnas generadas o vista):
- `TotalPercepciones` = Σ de las 6 percepciones.
- `ISR_Estimado`, `IMSS_Estimado` = estimados (histórico o % config, **no fiscal**).
- `NetoEstimado` = percepciones − ISR/IMSS est − otras ded. **Se congela al conciliar.**
- `Diferencia` = NetoReal − NetoEstimado (solo si hay CFDI).

Columnas **del CFDI** (las llena el timbrado, no se capturan): `ISR_Real` (deducción 002), `IMSS_Real` (001), `NetoReal` **[SENSIBLE]** (manda sobre el estimado si hay UUID), `UUID_CFDI` (su presencia = fila conciliada; puede concatenar varios), `TotalGravado`, `TotalExento`, `TotalRetenciones` (=ISR_Real+IMSS_Real), `TotalOtrasDeducciones` (=totalDeducciones−retenciones, piso 0).

Otras: `Pagado` boolean, `FechaPago` date, `ActualizadoEn` timestamptz.

> 🔥 `NetoEstimado`/`NetoReal`: cuidado con la regla de precedencia — **NetoReal manda si hay UUID**. Modelar como columna generada:
> `neto_final := CASE WHEN uuid_cfdi <> '' THEN neto_real ELSE neto_estimado END`.

#### 2.32 `nomina_captura_det` — líneas de concepto CFDI · CALIENTE

Sin PK natural → surrogate. Se **reemplaza en bloque** por empleado×periodo al conciliar.

`PeriodoID` FK, `EmpleadoID` FK, `Grupo` (percepcion\|deduccion\|otroPago), `Clave` (clave SAT), `Concepto`, `Importe`, `Gravado` (exacto del XML), `Exento` (exacto del XML), `Origen` ('CFDI'), `ActualizadoEn`.

- **Unicidad práctica:** `(PeriodoID, EmpleadoID, Grupo, Clave)`. **Índice** `(PeriodoID, EmpleadoID)`. Miles/año.

#### 2.33 `nomina_bonos`

`BonoID` PK ('BON-<epoch>-<rand>'), `Periodo` (YYYY-MM mensual, **no** el PeriodoID), `NumEmpleado` FK, `Nombre` **[PII]** denormalizado, `Concepto` (default 'Bono'), y (por el corte del mapeo, inferido) monto/usuario/timestamp. **Verificar el esquema real de esta hoja antes de migrar** — el mapeo entregado se truncó aquí.

#### 2.34 SBC bimestral (tablero derivado)

Mencionado en el mapeo/memoria: tablero de SBC por bimestre que **copia SBC** de `nomina_empleados`. Tratar como **vista** (no captura duplicada) o tabla de snapshot bimestral si se requiere congelar el histórico de cotización.

---

## 3. Seguridad (RLS)

El enmascarado que hoy es de **pantalla** (backend Apps Script decide qué mandar según rol) se mueve a la **base de datos**. Con RLS activa, ni la API ni una consulta directa pueden saltárselo.

### 3.1 Modelo de identidad y permisos

```sql
-- roles y permisos (reemplaza la hoja Roles y los flags como ver_datos_sensibles)
CREATE TABLE roles       (id, tenant_id, nombre, es_socio boolean, es_paciente boolean);
CREATE TABLE permisos    (id, clave);          -- ver_datos_sensibles, editar_egresos, ...
CREATE TABLE rol_permiso (rol_id, permiso_id);
CREATE TABLE usuarios    (id, tenant_id, email, rol_id, auth_user_id uuid);  -- ligado a auth.uid()
```

Función helper para políticas:

```sql
CREATE FUNCTION current_user_has(perm text) RETURNS boolean
LANGUAGE sql STABLE SECURITY DEFINER AS $$
  SELECT EXISTS (
    SELECT 1 FROM usuarios u
    JOIN rol_permiso rp ON rp.rol_id = u.rol_id
    JOIN permisos p ON p.id = rp.permiso_id
    WHERE u.auth_user_id = auth.uid() AND p.clave = perm);
$$;

CREATE FUNCTION current_tenant() RETURNS uuid
LANGUAGE sql STABLE AS $$
  SELECT tenant_id FROM usuarios WHERE auth_user_id = auth.uid();
$$;
```

### 3.2 Políticas por tabla (staff)

Toda tabla lleva **primero** el filtro de tenant, **luego** el de permiso:

```sql
ALTER TABLE bd_ingresos ENABLE ROW LEVEL SECURITY;

-- staff del tenant puede leer
CREATE POLICY staff_read ON bd_ingresos FOR SELECT
  USING (tenant_id = current_tenant());

-- pero los datos sensibles se enmascaran por VISTA, no por columna cruda:
REVOKE SELECT ON bd_ingresos FROM anon, authenticated;  -- nadie lee la tabla cruda
GRANT SELECT ON v_bd_ingresos TO authenticated;         -- todos leen la vista enmascarada
```

**Vista enmascarada** (el enmascarado de `_privPaciente`/`_egMask`/`_privRef` movido a SQL):

```sql
CREATE VIEW v_bd_ingresos WITH (security_invoker=true) AS
SELECT id, op, linea, fecha, categoria, producto, pvp, total_pagar, pagado, ...,
  CASE WHEN current_user_has('ver_datos_sensibles')
       THEN paciente_nombre ELSE 'OP-' || op END          AS paciente_nombre,
  CASE WHEN current_user_has('ver_datos_sensibles')
       THEN razon_social ELSE NULL END                    AS razon_social,
  CASE WHEN current_user_has('ver_datos_sensibles')
       THEN factura_rfc ELSE NULL END                     AS factura_rfc
FROM bd_ingresos
WHERE tenant_id = current_tenant();
```

Aplicar el mismo patrón a: `egresos` (enmascara proveedor/concepto/notas/factura_rfc/razon_social para rol **Socio**), `movimiento_bancario` y `caja_chica_movimiento` (referencia/observaciones), `cuentas_cobrar`/`abonos_cobrar`/`suscripciones_*`/`creditos_favor`/`creditos_consumo` (paciente/email), `pacientes` (email/RFC/fiscales), `nomina_empleados` (RFC/CURP/NSS/CLABE/salario).

### 3.3 Portal de pacientes — solo SUS filas

El paciente autenticado (rol `es_paciente`) ve exclusivamente lo suyo, vía la tabla puente `paciente_auth`:

```sql
CREATE POLICY paciente_solo_lo_suyo ON bd_ingresos FOR SELECT
USING (
  tenant_id = current_tenant()
  AND paciente_id IN (
    SELECT paciente_id FROM paciente_auth WHERE auth_user_id = auth.uid()
  )
);
```

- El portal **nunca** expone tablas internas (egresos, bancos, nómina, comisiones). Se le da una **vista dedicada** `v_portal_estado_cuenta` que solo une `cuentas_cobrar` + `abonos_cobrar` + `suscripciones_crio` de su `paciente_id`, ya enmascarada.
- El paciente **no escribe** en tablas de dinero: si el portal permite pagar, va por RPC controlado que valida el `paciente_id` contra `auth.uid()`.

### 3.4 PII cifrada a nivel columna

CURP, NSS, CLABE, RFC de empleados y fiscales de pacientes: cifrado con **Supabase Vault** / `pgcrypto`, descifrado solo en vista para roles con permiso. La `anon key` del frontend **jamás** ve el texto claro.

### 3.5 Auditoría (append-only)

Generalizar `bd_auditoria` a una tabla global `auditoria(tenant_id, timestamp, usuario, modulo, accion, entidad, entidad_id, campo, anterior, nuevo)`:
- **Append-only**: `GRANT INSERT` pero `REVOKE UPDATE, DELETE` a todos los roles de aplicación.
- Escrita por **triggers** en las tablas de dinero (no confiar en que la app la llene).
- Partición mensual + retención (§5.4).

### 3.6 Semántica PATCH obligatoria (bug de blanqueo)

🔥 **Bug conocido y recurrente (3+ incidentes):** un update con `campo || ''` **borra** lo que el payload no trajo. En Postgres:
- Las RPC de update usan `COALESCE(nuevo, actual)` o `UPDATE ... SET x = COALESCE($x, x)` — **solo campos presentes**.
- **Prohibido** el `UPDATE` de fila completa con nulls. Nunca borrar un dato que no se pueda re-crear (Alias, forma de pago→banco ya causaron pérdidas).

---

## 4. Multi-tenant para vender

### 4.1 Dos estrategias, decisión por cliente

| Estrategia | Aislamiento | Cuándo |
|---|---|---|
| **`tenant_id` + RLS (pooled)** | Lógico (una BD, filas etiquetadas) | Muchos clientes pequeños; onboarding barato; el default. |
| **Una base por cliente (isolated)** | Físico (BD/proyecto Supabase separado) | Cliente que exige aislamiento fuerte (salud, auditoría), o de alto volumen que quiere sus propias particiones/recursos. |

- **Empezar con `tenant_id` en TODAS las tablas** (ya está en §2.0). Migrar un tenant a base propia después es un `pg_dump` filtrado por `tenant_id` — barato **solo si** el `tenant_id` existió desde el día uno. Por eso se pone desde ahora aunque hoy haya un solo cliente.
- El `tenant_id` es la **primera columna de todo índice compuesto** — así el planner poda por tenant antes que nada.

### 4.2 Núcleo reusable vs específico de fertilidad

| Núcleo reusable (cualquier giro) | Específico de fertilidad (módulo desactivable) |
|---|---|
| Facturación / CFDI 4.0 (MX) | `lab_art`, `lab_fet` (ciclos ART / transferencias) |
| Bancos / conciliación / caja chica | `lab_inventario_crio` (banco criogénico) |
| CxP (egresos) + abonos + créditos proveedor | `suscripciones_crio` (almacenamiento crío) |
| CxC (cuentas por cobrar) + abonos + notas de crédito | `med_estimulacion` (protocolos de fármacos) |
| Inventario (catálogo único + ledger + combos + OC) | Campos clínicos de `pacientes` (médico tratante, ciclo) |
| Nómina MX (ISR/IMSS/SBC/CFDI nómina) | Cobranza Motor B (año gratis Hestia) |
| Catálogo de productos + listas de precio | Reportes tipo "Carta de Seguro" reembolso |
| Comisiones a vendedores / agencias / médicos externos | |

- Lo específico de fertilidad se aísla en un **schema `clinico`** (o tablas con prefijo) que un tenant de otro giro simplemente **no instala**. El núcleo vive en `core`.
- **Clonar para otro giro** = provisionar tenant nuevo con el schema `core` + las listas de precio / catálogo del giro, sin el schema `clinico`. La UI ya lee el menú de una tabla `menu` administrable → se recorta por tenant sin tocar código.

### 4.3 Config que hoy es Script Property → tablas de config por tenant

`COMISIONES_CFG`, `COBRANZA_CONFIG`, `COBRANZA_PAC_CFG`, `COBRANZA_DESC_CFG`, tarifas crío, formas de pago: mover a tablas `config_*` **con `tenant_id`** (o un key-value `tenant_config(tenant_id, clave, valor jsonb)`). `AUTH_SECRET` va a variables de entorno/Vault, **no** a tabla. Así cada cliente tiene su propia configuración sin tocar código ni properties globales.

---

## 5. Escalabilidad a 1M+ movimientos/año

### 5.1 Tablas calientes y su estrategia

| Tabla | Volumen | Partición | Índices clave |
|---|---|---|---|
| `bd_ingresos` | miles–1M+/año | RANGE por `fecha` (año) | `(tenant,op)`, `(tenant,fecha)`, `(tenant,paciente_id)`, `(tenant,origen_externo)` |
| `egresos` | cientos–miles/año | RANGE por `fecha` (año) | `(tenant,proveedor)`, `(tenant,fecha)`, `(tenant,pagado)`, `(tenant,estatus_flujo)` |
| `movimiento_bancario` | 1M+/año (MP) | RANGE por `fecha` + o LIST por `cuenta_key` | `(tenant,cuenta_key,seq)`, `(tenant,cuenta_key,fecha)`, `(tenant,mes)` |
| `movimientos_inventario` | alto | RANGE por `Fecha` | `(tenant,SKU,Fecha)` |
| `nomina_captura_det` | miles/año | — | `(PeriodoID,EmpleadoID)` |
| `bd_auditoria`/`auditoria` | crece sin techo | RANGE por `timestamp` (mes) | `(tenant,entidad,entidad_id)` |

### 5.2 Nada de "devolver TODO"

El código actual lee hojas completas (y ya topa en 500 filas de inventario). En Postgres, **toda lista de tabla caliente es paginada por keyset**:

```sql
-- página siguiente estable, sin OFFSET (que se degrada con el volumen)
SELECT ... FROM bd_ingresos
WHERE tenant_id = $t AND (fecha, id) < ($cursor_fecha, $cursor_id)
ORDER BY fecha DESC, id DESC
LIMIT 100;
```

### 5.3 Saldos y reportes materializados

- **Saldos de banco/caja/inventario:** el saldo corrido (§2.16.1, §2.17, §2.19.1) es el mayor riesgo de rendimiento. Para consultas de "saldo actual" muy frecuentes → **columna materializada por trigger** manteniendo `seq`. Para reportes de historia → **vista con window function**.
- **KPIs de dashboard** (totalAnual, ticketPromedio, topCategorías, aging 0-30/31-60/61-90/+90, comisiones MP por mes, Estado de Resultados, Board Deck): hoy se calculan al vuelo en cada request. → **vistas materializadas** refrescadas por `pg_cron` (p.ej. cada 5-15 min o al cierre de día), no en cada carga del dashboard.
- **Saldos de CxP / créditos:** vistas `v_cxp_saldos`, `v_creditos_proveedor_disponible`, `v_cuentas_cobrar_saldo` (§7). Si pesan, materializar.

### 5.4 Archivado de años viejos

Con partición por año, los años cerrados (2024, 2025) se `DETACH PARTITION` a tablespace frío / read-only. El dashboard consulta por defecto el año en curso; los históricos se piden explícitamente. La auditoría vieja se archiva/comprime por su partición mensual.

### 5.5 Connection pooling

Supabase incluye **PgBouncer** (modo transaction) — el frontend estático + PostgREST no abren conexiones directas por request. Las Edge Functions y jobs usan el pooler. Evita el agotamiento de conexiones a 1M+ operaciones.

---

## 6. Plan de migración por estrangulamiento (strangler fig)

**Sin big-bang.** Postgres corre **al lado** de Sheets; cada módulo se migra cuando su copia es **byte-idéntica** y verificada. Orden de menor a mayor riesgo de dinero:

| Fase | Módulo | Por qué en este orden |
|---|---|---|
| **0** | Infra: crear tenant Hestia, `tenants`, `usuarios`, `roles`, `permisos`, RLS base, auditoría | Sin tocar producción. |
| **1** | **Catálogos fríos**: `bd_listas`, `bd_productos`, `bd_precios`, `combos`, `config_cuentas`, `config_formas_pago` | Bajo volumen, sin dinero en movimiento, FKs raíz de todo lo demás. |
| **2** | **Pacientes** | Raíz de las FKs de cobranza; permite validar el enmascarado RLS y el portal **antes** de meter dinero. |
| **3** | **Inventario**: `movimientos_inventario`, `ordenes_compra(_lineas)` + invariante StockActual | Ledger con saldo materializado; ensaya el patrón trigger de saldo en una tabla menos crítica que bancos. |
| **4** | **Ingresos (`bd_ingresos`)** en modo **solo-lectura/espejo**: escribir en Sheets **y** en Postgres, comparar diario | La tabla más caliente; el espejo detecta divergencias sin arriesgar. Validar KPIs contra Sheets (byte-idéntico). |
| **5** | **CxC**: `cuentas_cobrar`, `abonos_cobrar`, `suscripciones_crio/mp`, `creditos_favor/consumo` | Depende de ingresos + pacientes ya estables. |
| **6** | **Egresos + CxP**: `egresos` (fila unificada), `abonos_cxp`, `creditos_proveedor`, `comisiones_generadas` | Dinero saliente; requiere el saldo-por-vista bien probado. |
| **7** | **Bancos y caja** (`movimiento_bancario`, `caja_chica_movimiento`) **al final** | El saldo corrido + conciliación son lo más delicado; migrar cuando todo lo que los alimenta ya está en Postgres. |
| **8** | **Nómina** (libro propio) | Aislado; PII máxima; se migra con cifrado de columnas ya en su sitio. |

### 6.1 Reglas del estrangulamiento

- **Ejecución en paralelo (dual-write) por módulo:** durante su fase, cada módulo escribe en Sheets y en Postgres; un job compara sumas de control (totales por mes, conteos, saldos) y alerta divergencias. Solo se corta Sheets cuando N días consecutivos cuadran **byte-idéntico**.
- **Byte-idéntico primero, refactor después:** replicar la semántica exacta (incluido `pagado` nulo = pagado, `fecha` vacía = CxP, saldo por orden de captura). Optimizar recién cuando la paridad esté probada.
- **Migrar por NOMBRE de columna, no por orden** (las columnas dinámicas append-safe no tienen posición garantizada).
- **Conciliación banco↔egreso/ingreso NO por etiqueta:** 🔥 `[Egreso #id]` solo existe desde 30-jun-2026; ene–may es carga histórica sin etiqueta. Cruzar por **fecha+monto (±2 días, N:1, valor absoluto)**. La etiqueta casi duplicó $11.9M una vez.
- **Mercado Pago:** 🔥 el ERP **nunca ha leído** de MP — **escribe** él las filas. Importar el reporte oficial de MP **duplicaría** ingresos. La comisión MP genera 1 egreso automático/mes (`MPCOM-YYYY-MM`), recalculado solo.
- **El año equivocado deja de existir:** al unir los libros por año en una tabla particionada, se elimina el riesgo de `_ingIdDeAnio`/`_egIdDeAnio` cayendo en silencio al 2026.

### 6.2 Validaciones de paridad obligatorias antes de cortar cada módulo

- Conteo de filas por mes y por año == Sheets.
- Suma de `total_pagar`, `egresos`, saldos de banco (última fila con saldo≠0) == Sheets.
- Aging de cobranza y saldos de CxP == lo que hoy calcula el ERP.
- Descartadas las filas placeholder (MP/Caja Chica) y las fechas re-normalizadas (DD/MM vs MM/DD de Caja Chica).

---

## 7. Vistas / columnas generadas / triggers a construir (resumen accionable)

| Hoy fórmula/recálculo | En Postgres |
|---|---|
| Saldo corrido banco (bancaria/credito/tpv) | Vista window `SUM OVER (PARTITION BY cuenta_key ORDER BY seq)` + opción materializada por trigger. Persistir `seq`. |
| Saldo disponible MP (liberado) vs por-liberar | Vista que filtra `liberado=TRUE`. |
| `pct_comision` MP | Columna generada `1 - total_venta/NULLIF(cobro,0)`. |
| Total corrido Caja Chica | Vista window arrancando de `es_remanente`. |
| `egresos.estatus_flujo` (CxP vs pagado vs cancelada) | Columna generada STORED. |
| Saldo de CxP | Vista `v_cxp_saldos = egresos − Σ abonos_cxp(no reversado)`. |
| `creditos_proveedor.monto_disponible` | Vista/trigger `monto − Σ consumos`. Nunca regenerar a monto original. |
| `creditos_favor.monto_credito` | Saldo vivo por upsert; modelar como `monto − Σ consumos`. |
| `cuentas_cobrar.monto_cargo` (auto-ingreso) | Trigger `AFTER INSERT/UPDATE ON bd_ingresos` = Facturado − COALESCE(pagado,total) por OP. |
| `StockActual` | Trigger `AFTER INSERT ON movimientos_inventario` (misma transacción, con lock). |
| Nómina: TotalPercepciones, Netos, Diferencia, precedencia Real/Estimado | Columnas generadas / vista de recibo. |
| KPIs, aging, Estado de Resultados, Board Deck, comisiones por mes | Vistas materializadas refrescadas por `pg_cron`. |
| Auditoría | Triggers append-only en tablas de dinero. |

---

## 8. Checklist "qué preparar YA sin tocar producción" (no-regret)

1. **Crear el proyecto Supabase** y el schema `core` + `clinico` vacíos, con `tenant_id` en todas las tablas desde el minuto uno.
2. **Modelar `tenants`, `usuarios`, `roles`, `permisos`, `rol_permiso`** y sembrar el tenant Hestia + los permisos que hoy son flags (`ver_datos_sensibles`, etc.) leyendo la **hoja Roles VIVA** (no la lista teórica — roles sin usuarios ≠ roles activos).
3. **DDL de catálogos fríos** (Fase 1) y cargar una copia **read-only** desde los libros de Productos/Listas/Cuentas para validar tipos y FKs por nombre. Reportar los `producto`/`SKU`/paciente que no emparejan.
4. **Escribir las vistas enmascaradas** (`v_bd_ingresos`, `v_egresos`, `v_pacientes`, …) y probar RLS con un usuario staff y un usuario paciente de prueba, **antes** de cargar dinero.
5. **Prototipar el saldo corrido** (window + `seq`) sobre una copia de una pestaña bancaria, y verificar que reproduce el saldo que hoy lee el ERP (última fila con saldo≠0).
6. **Job de dual-write/comparación** (esqueleto): script que lee Sheets y Postgres del mismo módulo y reporta divergencias de conteo/suma. Es la red de seguridad de todo el plan.
7. **Cifrado de columnas PII** (CURP/NSS/CLABE/RFC) configurado en Vault/pgcrypto y probado en `nomina_empleados` de prueba.
8. **Normalizadores de carga**: fechas Caja Chica (DD/MM vs MM/DD por >12), montos con `$`/`,`, `Sí/No`→boolean, `TRUE`/'TRUE'→boolean, `pagado` vacío→NULL (no 0), descarte de filas placeholder.
9. **Hosting**: publicar el frontend en Cloudflare Pages con `app.hestiafertility.com` (resuelve el bloqueo de github.io) apuntando aún a Apps Script — así la migración de backend no toca el hosting.
10. **Congelar el mapeo de nómina faltante**: el mapeo entregado se corta en `nomina_bonos`; extraer el esquema real de `Nomina_Bonos` y del tablero SBC antes de la Fase 8.

---

## 9. Riesgos y trampas — tarjeta de referencia rápida

1. `bd_ingresos.pagado` **vacío = pagado** (no 0). `COALESCE(pagado, total_pagar)`.
2. `egresos.fecha` **vacía = CxP**; con fecha = pagado. Una sola tabla.
3. Saldos de banco/caja/inventario son **corridos y se leen, no se capturan**. Persistir `seq`, ordenar por captura no por fecha.
4. `monto_disponible`/`monto_credito` son **saldos vivos**: regenerarlos duplica dinero.
5. `cuentas_cobrar.monto_cargo` (auto-ingreso) es **derivado de la venta**, no editable.
6. **Bug de blanqueo**: updates PATCH con `COALESCE`, nunca fila completa con nulls.
7. Columnas **dinámicas por encabezado**: migrar por nombre, no por orden.
8. Conciliación banco↔movimiento **por fecha+monto (±2d)**, no por la etiqueta `[Egreso #id]`.
9. **Mercado Pago no se importa** del reporte oficial (duplicaría): el ERP ya escribe esas filas.
10. **StockActual** solo lo escribe el ledger, en la misma transacción con lock.
11. **Nómina en libro propio** (NOMINA_SS_ID); la hoja Empleados de SHEET_ID es huérfana.
12. PII de paciente **prohibido** exportar a terceros: el enmascarado vive en la BD (RLS + vistas), no en pantalla.
13. `tenant_id` en TODO desde el día uno — es lo que hace vendible/clonable el sistema y barata la separación futura por cliente.

---

*Fin del documento. Construir en el orden de la §6, validando paridad byte-idéntica (§6.2) antes de cortar cada módulo de Sheets.*