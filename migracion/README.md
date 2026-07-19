# Migración VestaOS → base de datos real (cimiento)

Esto es el **cimiento** de la base nueva. **No toca el sistema actual** (Sheets + Apps Script sigue vivo y sirviéndose desde Cloudflare/githack). Son archivos SQL listos para correr en un proyecto Supabase nuevo. Ver el plano completo en [`../docs/ARQUITECTURA_ERP.md`](../docs/ARQUITECTURA_ERP.md).

## Qué hay aquí

| Archivo | Qué crea |
|---|---|
| `01_core_schema.sql` | Esquema núcleo: tenants (multi-cliente), usuarios (ligados a la auth de Supabase), roles, permisos, auditoría. |
| `02_clinico_pacientes.sql` | Tabla `pacientes` (el primer dominio a migrar) + enlace para el portal de pacientes. |
| `03_funciones_rls.sql` | **La seguridad**: funciones de permiso + Row-Level Security + la vista enmascarada. El enmascarado deja de ser de pantalla y lo impone la base. |
| `04_seed.sql` | Siembra el catálogo de permisos + el tenant Hestia con sus roles (idéntico al ERP de hoy). |

## Cómo correrlo (≈15 min, sin riesgo)

1. Crea un proyecto en **supabase.com** (plan gratis alcanza para empezar). Guarda la contraseña de la base.
2. Abre **SQL Editor** en el panel de Supabase.
3. Pega y corre **en orden**: `01` → `02` → `03` → `04`. Cada uno debe terminar sin error.
4. **Exponer los esquemas a la API** (si no, la app no ve nada): Supabase → **Settings → API → Exposed schemas** → agrega `core` y `clinico`. (Los `grant` necesarios ya están en `03`.)
5. Verifica: en **Table Editor** deben aparecer los esquemas `core` y `clinico` con sus tablas; en `core.permisos` los permisos; en `core.roles` los 7 roles del tenant Hestia.

## Cómo se prueba la seguridad (antes de meter datos reales)

En el SQL Editor puedes simular ser un usuario y ver que la RLS funciona:
- Crea 2 usuarios de prueba en **Authentication** (uno con rol `captura`, otro `socio`) y enlázalos con la plantilla del final de `04_seed.sql`.
- Como `captura` → `select * from clinico.v_pacientes` muestra correo/RFC.
- Como `socio` → la MISMA consulta muestra esos campos en `null` (enmascarado por la base, no por el frontend).
- Como un usuario **paciente** (portal) → solo ve **su propia** fila.

## Cargar pacientes en modo espejo + red de seguridad

`migrar_pacientes.mjs` copia la hoja Pacientes a la base nueva **sin apagar nada** (el ERP viejo sigue vivo) y luego **compara Sheets vs Postgres campo por campo** para probar que cuadran. Solo Node 18+ (sin instalar dependencias).

1. **Exporta la hoja Pacientes a CSV**: en Google Sheets, con la hoja Pacientes activa → *Archivo → Descargar → CSV*.
2. **Pon los secretos en variables de entorno** (nunca en el código ni en el repo). La `service_role` está en Supabase → *Project Settings → API*. En PowerShell:
   ```powershell
   $env:SUPABASE_URL = "https://xxxx.supabase.co"
   $env:SUPABASE_SERVICE_KEY = "<tu service_role>"   # SECRETA — ignora la RLS
   ```
3. **Carga (upsert idempotente por folio)** y luego **verifica**:
   ```powershell
   node migrar_pacientes.mjs cargar    pacientes.csv
   node migrar_pacientes.mjs verificar pacientes.csv   # debe dar 0 diferencias
   ```

`verificar` imprime la **RED DE SEGURIDAD**: iguales / difieren / faltan / de más / sin ID. Correr `cargar` de nuevo es seguro (no duplica; re-actualiza por folio). Pacientes **sin ID** se avisan y se omiten (sin llave no hay upsert idempotente ni comparación). La comparación pagina, así que es fiable con más de 1000 pacientes.

## Lo que sigue (no en este cimiento)

- Resto de dominios por estrangulamiento: catálogos → ingresos/egresos → **bancos al final**.
- Portal de pacientes sobre este esquema.

## Notas importantes

- **Contraseñas:** ya no se guardan aquí — las maneja Supabase Auth (bcrypt). El problema de texto plano del sistema viejo desaparece en la base nueva.
- **Multi-tenant / vender:** todo lleva `tenant_id`. Para otro cliente/giro se crea otro tenant (o, para aislamiento total, otra base con este mismo esquema). El esquema `core` es reusable; `clinico` (fertilidad) es desactivable.
- **Semilla de roles:** `04_seed.sql` usa los defaults del código. Antes de producción, **conciliar con la hoja Roles viva** por si hay roles/permisos personalizados.
