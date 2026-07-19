-- ============================================================================
-- VestaOS ERP · Parche 06 · Correo de paciente NO único
-- ----------------------------------------------------------------------------
-- SOLO se necesita si ya corriste 02 ANTES de este arreglo. En una instalación
-- nueva, el archivo 02 ya crea el índice no-único y este parche sobra.
--
-- Síntoma que corrige:  REST 400  "duplicate key value violates unique
-- constraint pacientes_tenant_email_uniq" (23505) al cargar pacientes.
--
-- Motivo: en fertilidad las PAREJAS comparten un mismo correo, y el ERP de hoy
-- AVISA del correo repetido pero NO lo bloquea. La constraint dura de la BD era
-- más estricta que el ERP real e impedía el espejo fiel de la hoja. Se cambia
-- por un índice normal (sirve para búsqueda y para el aviso de antiduplicado,
-- que vive en la capa de aplicación).
--
-- Idempotente. Pégalo en Supabase → SQL Editor → Run.
-- ============================================================================

drop index if exists clinico.pacientes_tenant_email_uniq;

create index if not exists pacientes_tenant_email
  on clinico.pacientes (tenant_id, lower(email))
  where email is not null and email <> '' and deleted_at is null;
