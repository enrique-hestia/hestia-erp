-- ============================================================================
-- VestaOS ERP · Parche 05 · GRANTS para service_role (script de migración)
-- ----------------------------------------------------------------------------
-- SOLO se necesita si ya corriste 01–04 ANTES de este arreglo. En una instalación
-- nueva, el archivo 03 ya incluye estos grants y este parche sobra.
--
-- Síntoma que corrige:  REST 403  "permission denied for schema core" (42501)
-- al correr migrar_pacientes.mjs.
--
-- Causa: la llave `service_role` (la SECRETA del backend) ignora la RLS pero NO
-- los grants, y Supabase no le da acceso automático a esquemas creados a mano
-- (core, clinico). Aquí se lo concedemos. service_role es la llave de servidor:
-- acceso total es correcto y esperado (se guarda en secreto, nunca en el navegador).
--
-- Idempotente: correr de nuevo no hace daño.
-- Pégalo en Supabase → SQL Editor → Run.
-- ============================================================================

grant usage on schema core, clinico to service_role;
grant all on all tables    in schema core, clinico to service_role;
grant all on all sequences in schema core, clinico to service_role;

-- Para las tablas/secuencias que se agreguen después (ingresos, egresos, …)
alter default privileges in schema core, clinico grant all on tables    to service_role;
alter default privileges in schema core, clinico grant all on sequences to service_role;

-- Comprobación rápida (debe listar core y clinico con has_schema_privilege = true):
select nspname as esquema,
       has_schema_privilege('service_role', nspname, 'USAGE') as service_role_usage
from pg_namespace
where nspname in ('core','clinico')
order by nspname;
