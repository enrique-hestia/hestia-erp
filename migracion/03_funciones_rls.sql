-- ============================================================================
-- VestaOS ERP · Archivo 03 de 4 · FUNCIONES DE SEGURIDAD + RLS
-- ----------------------------------------------------------------------------
-- El corazón de la seguridad. Aquí los permisos y el enmascarado dejan de vivir
-- en el frontend (que se salta con las herramientas del navegador) y pasan a la
-- BASE: aunque alguien llame la API directo, la base decide qué filas y qué
-- columnas devuelve. `auth.uid()` es el usuario autenticado (Supabase Auth).
--
-- Las funciones son SECURITY DEFINER: corren como dueño (postgres) para poder
-- leer las tablas de roles SIN disparar la RLS de esas mismas tablas (evita
-- recursión infinita en las políticas). Son STABLE (mismo resultado dentro de
-- la consulta) para que el planificador las cachee.
-- ============================================================================

-- ¿A qué tenant pertenece el usuario actual? Staff (core.usuarios) o, si no,
-- paciente del portal (clinico.pacientes.auth_user_id).
create or replace function core.current_tenant()
returns uuid language sql stable security definer set search_path = '' as $$
  select coalesce(
    (select tenant_id from core.usuarios      where id = auth.uid()),
    (select tenant_id from clinico.pacientes  where auth_user_id = auth.uid() limit 1)
  )
$$;

-- ¿El usuario actual es staff (personal del cliente)?
create or replace function core.is_staff()
returns boolean language sql stable security definer set search_path = '' as $$
  select exists (select 1 from core.usuarios where id = auth.uid() and activo)
$$;

-- ¿Su rol es "super" (admin/director) → ve TODO, bypass de permisos?
create or replace function core.is_super()
returns boolean language sql stable security definer set search_path = '' as $$
  select exists (
    select 1 from core.usuarios u
    join core.roles r on r.tenant_id = u.tenant_id and r.clave = u.rol_clave
    where u.id = auth.uid() and u.activo and r.es_super
  )
$$;

-- ¿El usuario actual tiene el permiso `perm`? (super = sí a todo; '*' = todo)
create or replace function core.has_permission(perm text)
returns boolean language sql stable security definer set search_path = '' as $$
  select core.is_super() or exists (
    select 1 from core.usuarios u
    join core.rol_permiso rp on rp.tenant_id = u.tenant_id and rp.rol_clave = u.rol_clave
    where u.id = auth.uid() and u.activo and rp.permiso_clave in (perm, '*')
  )
$$;

-- ¿Puede ver datos sensibles de pacientes/empleados? (el eje del enmascarado)
create or replace function core.can_see_sensitive()
returns boolean language sql stable security definer set search_path = '' as $$
  select core.has_permission('ver_datos_sensibles')
$$;

-- id de paciente del portal (la fila propia del paciente autenticado), o null.
create or replace function core.current_patient_id()
returns bigint language sql stable security definer set search_path = '' as $$
  select id from clinico.pacientes where auth_user_id = auth.uid() and deleted_at is null limit 1
$$;

-- Escribe en la bitácora con el actor REAL (auth.uid), no uno que el cliente
-- invente. Es la ÚNICA vía de insert (auditoria no tiene policy de INSERT).
create or replace function core.audit(_accion text, _entidad text default null, _entidad_id text default null, _detalle jsonb default null)
returns void language sql security definer set search_path = '' as $$
  insert into core.auditoria (tenant_id, actor, actor_email, accion, entidad, entidad_id, detalle)
  select core.current_tenant(), auth.uid(),
         (select email from core.usuarios where id = auth.uid()),
         _accion, _entidad, _entidad_id, _detalle
  where core.current_tenant() is not null;
$$;

-- Estas funciones NO deben ser ejecutables por anónimos (endurecimiento).
-- Nota de dueño: en Supabase el editor SQL corre como `postgres` (BYPASSRLS), y
-- por eso estas funciones SECURITY DEFINER leen roles/usuarios sin disparar su
-- propia RLS (no hay recursión). Recrearlas bajo otro dueño SIN BYPASSRLS
-- rompería eso — mantenerlas propiedad de postgres.
revoke execute on function
  core.current_tenant(), core.is_staff(), core.is_super(),
  core.has_permission(text), core.can_see_sensitive(), core.current_patient_id(),
  core.audit(text, text, text, jsonb)
  from public;
grant execute on function
  core.current_tenant(), core.is_staff(), core.is_super(),
  core.has_permission(text), core.can_see_sensitive(), core.current_patient_id(),
  core.audit(text, text, text, jsonb)
  to authenticated;

-- ============================================================================
-- RLS · negar por default y abrir con políticas explícitas
-- ============================================================================

-- ── core.tenants ─────────────────────────────────────────────────────────
-- OJO: current_tenant() también resuelve para un PACIENTE del portal, así que
-- estas lecturas exigen is_staff() — si no, un paciente leería core.tenants.config
-- (tarifas/parámetros) y la matriz de roles/permisos de la clínica.
alter table core.tenants enable row level security;
create policy tenants_read on core.tenants for select
  using (id = core.current_tenant() and core.is_staff());

-- ── core.usuarios ────────────────────────────────────────────────────────
alter table core.usuarios enable row level security;
create policy usuarios_read on core.usuarios for select
  using (tenant_id = core.current_tenant() and core.is_staff());
create policy usuarios_admin_write on core.usuarios for all
  using      (tenant_id = core.current_tenant() and core.is_super())
  with check (tenant_id = core.current_tenant() and core.is_super());

-- ── core.roles / permisos / rol_permiso ──────────────────────────────────
alter table core.roles       enable row level security;
alter table core.permisos    enable row level security;
alter table core.rol_permiso enable row level security;
create policy roles_read on core.roles for select using (tenant_id = core.current_tenant() and core.is_staff());
create policy roles_admin_write on core.roles for all
  using (tenant_id = core.current_tenant() and core.is_super())
  with check (tenant_id = core.current_tenant() and core.is_super());
create policy permisos_read on core.permisos for select using (auth.uid() is not null);   -- catálogo, solo logueados
create policy rolperm_read on core.rol_permiso for select using (tenant_id = core.current_tenant() and core.is_staff());
create policy rolperm_admin_write on core.rol_permiso for all
  using (tenant_id = core.current_tenant() and core.is_super())
  with check (tenant_id = core.current_tenant() and core.is_super());

-- ── core.auditoria — append-only ─────────────────────────────────────────
alter table core.auditoria enable row level security;
-- SIN policy de INSERT: nadie escribe directo (evita que un cliente falsifique
-- el actor). Se registra vía core.audit(), que fija actor = auth.uid(). Sin
-- policies de UPDATE/DELETE → la bitácora es inmutable.
create policy auditoria_read on core.auditoria for select
  using (tenant_id = core.current_tenant() and core.has_permission('sheets_access'));

-- ── clinico.pacientes ────────────────────────────────────────────────────
-- Dos poblaciones distintas leen esta tabla:
--   · STAFF: ve a TODOS los pacientes de su tenant (recepción los necesita).
--   · PACIENTE (portal): ve SOLO su propia fila (auth_user_id = auth.uid()).
alter table clinico.pacientes enable row level security;
-- SIN policy de SELECT aquí a propósito: la lectura va por la vista v_pacientes
-- (que además revoca el SELECT directo a esta tabla). Solo quedan escrituras,
-- e INSERT/UPDATE por separado — sin DELETE → el hard-delete queda DENEGADO por
-- RLS (el borrado es lógico, vía deleted_at con UPDATE).
create policy pacientes_ins on clinico.pacientes for insert
  with check (tenant_id = core.current_tenant() and core.has_permission('editar_pacientes'));
create policy pacientes_upd on clinico.pacientes for update
  using      (tenant_id = core.current_tenant() and core.has_permission('editar_pacientes'))
  with check (tenant_id = core.current_tenant() and core.has_permission('editar_pacientes'));

-- ============================================================================
-- VISTA ENMASCARADA · el enmascarado que hoy es de PANTALLA, ahora en la BASE.
-- ----------------------------------------------------------------------------
-- IMPORTANTE (hallazgo de auditoría): la vista es la ÚNICA puerta de lectura.
-- Se REVOCA el SELECT directo a clinico.pacientes; si se dejara leíble, un rol
-- sin 'ver_datos_sensibles' consultaría la tabla por PostgREST y vería el
-- correo/RFC SIN máscara (el case de la vista nunca correría). Por eso:
--   · La vista es security_invoker=FALSE → corre como dueño y trae los filtros
--     de tenant/portal EXPLÍCITOS en el WHERE (obligatorios: sin ellos fugaría
--     entre clientes).
--   · Se enmascara TAMBIÉN el nombre: socio/viewer ven folio/OP, no el nombre
--     (pilar de privacidad). El PACIENTE del portal ve SIEMPRE lo suyo completo.
-- ============================================================================
create or replace view clinico.v_pacientes
with (security_invoker = false) as
select
  id, tenant_id, folio, lista_precios,
  canal, origen, medico_tratante, pais, idioma,
  case when core.can_see_sensitive() or auth_user_id = auth.uid()
       then nombre else coalesce(folio, 'OP-'||id::text) end            as nombre,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then email               else null end as email,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then telefono            else null end as telefono,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then rfc                 else null end as rfc,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then razon_social        else null end as razon_social,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then codigo_postal       else null end as codigo_postal,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then uso_cfdi            else null end as uso_cfdi,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then regimen_fiscal      else null end as regimen_fiscal,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then forma_pago_habitual else null end as forma_pago_habitual,
  case when core.can_see_sensitive() or auth_user_id = auth.uid() then fecha_nacimiento    else null end as fecha_nacimiento,
  activo, created_at
from clinico.pacientes
where deleted_at is null
  and tenant_id = core.current_tenant()
  and (core.is_staff() or auth_user_id = auth.uid());   -- staff: su tenant · paciente: solo lo suyo

-- La tabla base NO se lee directo (solo por la vista); las escrituras del staff
-- siguen entrando por INSERT/UPDATE (gateadas por las policies de arriba). La
-- carga de migración usa service_role, que ignora RLS y grants.
revoke all      on clinico.pacientes  from authenticated, anon;
grant  insert, update on clinico.pacientes to authenticated;   -- NO select
grant  select   on clinico.v_pacientes to authenticated;

comment on view clinico.v_pacientes is 'ÚNICA puerta de lectura de pacientes. Enmascara nombre/correo/RFC/fiscales para roles sin ver_datos_sensibles; el paciente del portal ve lo suyo. La tabla base tiene revocado el SELECT.';

-- ============================================================================
-- GRANTS de acceso para PostgREST (la API REST de Supabase).
-- La RLS de arriba sigue filtrando las FILAS; estos grants solo permiten que el
-- rol `authenticated` ALCANCE los objetos. Sin esto la app no lee nada (falla
-- cerrado, no es fuga). Se concede a `authenticated`, NO a `anon` (todo exige
-- login; con auth.uid() null las policies no devuelven nada de todos modos).
-- ⚠ ADEMÁS, en el panel de Supabase: Settings → API → Exposed schemas, agregar
--    `core` y `clinico` (o db-schemas = public,core,clinico) — si no, PostgREST
--    ni ve estos esquemas.
-- ============================================================================
grant usage on schema core, clinico to authenticated;
grant select on core.tenants, core.usuarios, core.roles, core.permisos, core.rol_permiso, core.auditoria to authenticated;
-- clinico.pacientes: NO se concede SELECT (se lee por v_pacientes, ya concedida);
-- insert/update ya concedidos arriba. La bitácora se escribe vía core.audit().
--
-- ⚠ NOTA de escritura (fase futura, no la fase espejo): como se revocó el SELECT
--    en clinico.pacientes, un INSERT/UPDATE con `Prefer: return=representation`
--    (PostgREST intenta leer la fila escrita) fallará y revertirá la escritura.
--    Al escribir pacientes, usar `Prefer: return=minimal`, o exponer la escritura
--    por una RPC security-definer que devuelva la fila enmascarada de v_pacientes.
