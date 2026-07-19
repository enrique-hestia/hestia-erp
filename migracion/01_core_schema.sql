-- ============================================================================
-- VestaOS ERP · Cimiento de la base nueva (Supabase / PostgreSQL)
-- Archivo 01 de 4 · ESQUEMA NÚCLEO (multi-tenant + seguridad)
-- ----------------------------------------------------------------------------
-- Correr en orden: 01 (este) → 02 (pacientes) → 03 (funciones+RLS) → 04 (seed).
-- Se corre en un proyecto Supabase NUEVO. No toca nada del sistema actual
-- (Sheets/Apps Script sigue vivo). Ver docs/ARQUITECTURA_ERP.md.
--
-- Idea central:
--   · Cada fila lleva `tenant_id` → un solo motor sirve a muchos clientes
--     (multi-tenant), y para vender a otro giro se crea otro tenant o, si se
--     quiere aislamiento total, otra base con este mismo esquema.
--   · La AUTENTICACIÓN la maneja Supabase Auth (tabla auth.users, contraseñas
--     con bcrypt) — así desaparece el problema de contraseñas en texto plano.
--   · Los PERMISOS y el enmascarado se aplican con Row-Level Security (RLS) en
--     la BASE, no en el frontend: aunque el cliente intente saltárselo, la base
--     se niega. Eso vive en el archivo 03.
-- ============================================================================

create extension if not exists pgcrypto;      -- gen_random_uuid()

create schema if not exists core;             -- núcleo reusable (cualquier giro)
create schema if not exists clinico;          -- específico de fertilidad (desactivable)

-- ----------------------------------------------------------------------------
-- TENANTS — cada cliente/empresa. La raíz de todo el aislamiento.
-- ----------------------------------------------------------------------------
create table core.tenants (
  id           uuid primary key default gen_random_uuid(),
  slug         text not null unique,               -- 'hestia', 'clienteX'
  nombre       text not null,
  giro         text,                               -- 'fertilidad', 'dental', etc.
  activo       boolean not null default true,
  config       jsonb not null default '{}'::jsonb, -- parámetros por cliente (ISN, tarifas, etc.)
  created_at   timestamptz not null default now()
);
comment on table core.tenants is 'Cada empresa/cliente del ERP. tenant_id de todas las tablas apunta aquí.';

-- ----------------------------------------------------------------------------
-- PERMISOS — catálogo de permisos operativos (mismos que hoy usa el ERP).
-- ----------------------------------------------------------------------------
create table core.permisos (
  clave       text primary key,                    -- 'ver_datos_sensibles', 'editar_egresos'
  label       text not null,
  categoria   text
);
comment on table core.permisos is 'Catálogo global de permisos. Se siembra en el archivo 04.';

-- ----------------------------------------------------------------------------
-- ROLES — por tenant. admin/director ven todo; el resto según sus permisos.
-- ----------------------------------------------------------------------------
create table core.roles (
  id          bigint generated always as identity primary key,
  tenant_id   uuid not null references core.tenants(id) on delete cascade,
  clave       text not null,                        -- 'admin','gerente','captura','socio'...
  nombre      text not null,
  es_super    boolean not null default false,       -- admin/director = ve TODO (bypass de permiso)
  solo_lectura boolean not null default false,
  created_at  timestamptz not null default now(),
  unique (tenant_id, clave)
);

-- ROL ↔ PERMISO (muchos a muchos). Usar clave '*' para "todos".
create table core.rol_permiso (
  tenant_id     uuid not null references core.tenants(id) on delete cascade,
  rol_clave     text not null,
  permiso_clave text not null,                      -- FK lógica a core.permisos.clave o '*'
  primary key (tenant_id, rol_clave, permiso_clave)
);

-- ----------------------------------------------------------------------------
-- USUARIOS (staff) — EXTIENDE auth.users de Supabase (que guarda la contraseña
-- hasheada con bcrypt). Aquí solo van tenant, rol y datos de presentación.
-- ----------------------------------------------------------------------------
create table core.usuarios (
  id          uuid primary key references auth.users(id) on delete cascade,
  tenant_id   uuid not null references core.tenants(id) on delete cascade,
  email       text not null,
  nombre      text,
  alias       text,                                 -- lo que se ve en pantalla/chat
  rol_clave   text not null default 'viewer',
  activo      boolean not null default true,
  created_at  timestamptz not null default now(),
  unique (tenant_id, email)
);
comment on table core.usuarios is 'Personal del cliente. La contraseña NO vive aquí: la maneja Supabase Auth (auth.users, bcrypt).';

create index on core.usuarios (tenant_id);
create index on core.roles (tenant_id);
create index on core.rol_permiso (tenant_id, rol_clave);

-- ----------------------------------------------------------------------------
-- AUDITORÍA — append-only. Quién hizo qué y cuándo (base para el portal ARCO
-- y para investigar incidentes). Particionable por mes cuando crezca.
-- ----------------------------------------------------------------------------
create table core.auditoria (
  id          bigint generated always as identity primary key,
  tenant_id   uuid not null references core.tenants(id) on delete cascade,
  ts          timestamptz not null default now(),
  actor       uuid,                                 -- auth.uid() del que actuó
  actor_email text,
  accion      text not null,                        -- 'login','ver_paciente','editar_egreso'...
  entidad     text,
  entidad_id  text,
  detalle     jsonb
);
create index on core.auditoria (tenant_id, entidad, entidad_id);
create index on core.auditoria (tenant_id, ts desc);
comment on table core.auditoria is 'Bitácora inmutable. Nunca se hace UPDATE/DELETE (solo INSERT).';
