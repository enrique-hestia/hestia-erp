-- ============================================================================
-- VestaOS ERP · PRUEBA DE SEGURIDAD (RLS + enmascarado impuesto por la BD)
-- ----------------------------------------------------------------------------
-- Demuestra que el candado lo pone la BASE DE DATOS, no el frontend: aunque
-- alguien llame la API directo, la base decide qué FILAS y qué COLUMNAS entrega.
--
-- CÓMO USARLO: el SQL Editor solo muestra el resultado del ÚLTIMO SELECT, así que
-- corre UN BLOQUE A LA VEZ (selecciónalo con el mouse y pulsa Run / Ctrl+Enter).
--
-- ANTES: crea 2 usuarios de prueba en Supabase → Authentication → Users →
--   "Add user" → "Create new user" (marca "Auto Confirm User"):
--     · test-captura@hestia.local   (será rol 'captura' — SÍ ve datos sensibles)
--     · test-socio@hestia.local     (será rol 'socio'   — NO ve datos sensibles)
--   Copia el "User UID" (UUID) de cada uno y pégalos abajo.
-- ============================================================================


-- ██ BLOQUE 0 — enlazar los 2 usuarios de prueba (córrelo UNA vez) ██████████
-- Reemplaza los dos UUID por los que copiaste de Authentication.
insert into core.usuarios (id, tenant_id, email, nombre, alias, rol_clave, activo)
select '<<UUID_CAPTURA>>', t.id, 'test-captura@hestia.local', 'Prueba Captura', 'captura', 'captura', true
from core.tenants t where t.slug = 'hestia'
on conflict (id) do update set rol_clave = excluded.rol_clave, activo = true;

insert into core.usuarios (id, tenant_id, email, nombre, alias, rol_clave, activo)
select '<<UUID_SOCIO>>', t.id, 'test-socio@hestia.local', 'Prueba Socio', 'socio', 'socio', true
from core.tenants t where t.slug = 'hestia'
on conflict (id) do update set rol_clave = excluded.rol_clave, activo = true;


-- ██ BLOQUE A — como CAPTURA (tiene ver_datos_sensibles) → VE la PII ████████
-- Esperado: columnas email / rfc / telefono CON datos reales.
begin;
  select set_config('request.jwt.claims', '{"sub":"<<UUID_CAPTURA>>","role":"authenticated"}', true);
  set local role authenticated;
  select folio, nombre, email, rfc, telefono
  from clinico.v_pacientes
  order by folio
  limit 8;
rollback;


-- ██ BLOQUE B — como SOCIO (NO tiene ver_datos_sensibles) → BD ENMASCARA ████
-- Esperado (MISMA consulta, mismo usuario de datos): nombre = folio "OP-…",
-- y email / rfc / telefono en NULL. El enmascarado lo hizo la BASE, no la app.
begin;
  select set_config('request.jwt.claims', '{"sub":"<<UUID_SOCIO>>","role":"authenticated"}', true);
  set local role authenticated;
  select folio, nombre, email, rfc, telefono
  from clinico.v_pacientes
  order by folio
  limit 8;
rollback;


-- ██ BLOQUE C — ATAQUE: leer la TABLA BASE directo, saltándose la vista ██████
-- Así intentaría un atacante que llama la API a clinico.pacientes en vez de a la
-- vista. DEBE FALLAR con "permission denied for table pacientes". Ese error ES
-- la prueba: la única puerta de lectura es la vista enmascarada.
begin;
  set local role authenticated;
  select nombre, email, rfc from clinico.pacientes limit 1;   -- ← debe dar ERROR
rollback;


-- ██ BLOQUE D (opcional) — un PACIENTE del portal SOLO ve SU fila ████████████
-- 1) Crea en Authentication un usuario para una paciente y copia su UUID.
-- 2) Enlázalo a un folio REAL (usa uno que hayas visto en el BLOQUE A, ej HEC-015):
update clinico.pacientes set auth_user_id = '<<UUID_PACIENTE>>'
where tenant_id = (select id from core.tenants where slug = 'hestia')
  and folio = '<<FOLIO_REAL>>';
-- 3) Impersona a esa paciente. Esperado: UNA sola fila (la suya), con SU PII
--    visible (a su propia info sí tiene derecho), y NADA de las demás pacientes.
begin;
  select set_config('request.jwt.claims', '{"sub":"<<UUID_PACIENTE>>","role":"authenticated"}', true);
  set local role authenticated;
  select folio, nombre, email from clinico.v_pacientes;   -- ← solo su fila
rollback;


-- ██ LIMPIEZA (cuando termines de probar) ████████████████████████████████████
-- delete from core.usuarios where email in ('test-captura@hestia.local','test-socio@hestia.local');
-- update clinico.pacientes set auth_user_id = null where folio = '<<FOLIO_REAL>>';
-- (y borra los usuarios de prueba en Authentication → Users)
