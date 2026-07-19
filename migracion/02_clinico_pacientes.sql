-- ============================================================================
-- VestaOS ERP · Archivo 02 de 4 · PACIENTES (primer dominio a migrar)
-- ----------------------------------------------------------------------------
-- Pacientes es el mejor primer módulo: es un catálogo (bajo riesgo, no mueve
-- dinero) y es donde vive la PII más sensible, así que probar aquí la seguridad
-- de la BD vale para todo lo demás. Se migra en modo ESPEJO (solo lectura):
-- se carga una copia desde la hoja Pacientes y se compara contra Sheets con el
-- script de migración, sin apagar nada del sistema actual.
--
-- El campo `auth_user_id` es la llave del PORTAL DE PACIENTES: cuando una
-- paciente se registre en el portal, su cuenta de Supabase Auth se enlaza aquí,
-- y la RLS (archivo 03) hará que SOLO vea su propia fila.
-- ============================================================================

create table clinico.pacientes (
  id                  bigint generated always as identity primary key,
  tenant_id           uuid not null references core.tenants(id) on delete cascade,
  folio               text,                     -- HEC-001 (continuidad visual; no es la PK)
  nombre              text not null,            -- [SENSIBLE]
  fecha_nacimiento    date,
  email               text,                     -- [SENSIBLE]
  telefono            text,                     -- [SENSIBLE]
  origen              text,
  canal               text,                     -- Hestia / Externo (catálogo Orígenes)
  medico_tratante     text,                     -- [clínico] específico de fertilidad
  pais                text,
  idioma              text,                     -- idioma de sus documentos
  observaciones       text,
  -- Datos fiscales (CFDI) — [SENSIBLE]
  razon_social        text,
  rfc                 text,
  codigo_postal       text,
  uso_cfdi            text,
  regimen_fiscal      text,
  forma_pago_habitual text,
  lista_precios       text not null default 'General',
  -- Portal de pacientes: enlace a su cuenta de Supabase Auth (nullable hasta
  -- que se registre). La RLS del portal filtra por este campo.
  auth_user_id        uuid references auth.users(id) on delete set null,
  activo              boolean not null default true,
  created_at          timestamptz not null default now(),
  updated_at          timestamptz not null default now(),
  deleted_at          timestamptz,              -- soft-delete: nunca hard-delete
  unique (tenant_id, folio)
);

-- Antiduplicado por correo (lo que hoy valida el ERP a mano): un correo no se
-- repite dentro del mismo tenant. Índice parcial: ignora correos vacíos y
-- pacientes borrados.
create unique index pacientes_tenant_email_uniq
  on clinico.pacientes (tenant_id, lower(email))
  where email is not null and email <> '' and deleted_at is null;

create index on clinico.pacientes (tenant_id, lower(nombre));
create index on clinico.pacientes (auth_user_id);   -- lookups del portal

-- updated_at automático
create or replace function core.tg_touch_updated_at()
returns trigger language plpgsql set search_path = '' as $$
begin new.updated_at = now(); return new; end $$;

create trigger touch_pacientes before update on clinico.pacientes
  for each row execute function core.tg_touch_updated_at();

comment on table clinico.pacientes is 'Padrón único de pacientes. auth_user_id enlaza al portal. Campos [SENSIBLE] enmascarados por RLS/vista (archivo 03).';
