-- ============================================================================
-- VestaOS ERP · Archivo 04 de 4 · SEMILLA (permisos + tenant Hestia + roles)
-- ----------------------------------------------------------------------------
-- Siembra el catálogo de permisos (idéntico al del ERP actual) y el tenant
-- Hestia con sus roles y qué permisos tiene cada uno (tomado de
-- ROLE_PERMS_DEFAULT del frontend).
-- ⚠ VERIFICAR contra la hoja Roles VIVA antes de producción: si el cliente creó
--    roles a mano o cambió permisos, ajústalos aquí (roles sin usuarios ≠ roles
--    activos). Los usuarios (staff) se crean en Supabase Auth y se enlazan
--    abajo (plantilla comentada).
-- Idempotente: on conflict do nothing / do update.
-- ============================================================================

-- ── Catálogo de permisos ──────────────────────────────────────────────────
insert into core.permisos (clave, label, categoria) values
  ('bank_create',               'Crear bancos nuevos',                         'Bancos'),
  ('bank_capture',              'Capturar movimientos bancarios',              'Bancos'),
  ('bank_edit',                 'Editar movimientos bancarios',                'Bancos'),
  ('bank_delete',               'Eliminar movimientos bancarios',              'Bancos'),
  ('bank_liberado',             'Marcar liberado (MP)',                        'Bancos'),
  ('editar_ingresos',           'Editar operaciones de ingreso',               'Ingresos'),
  ('cancelar_ingresos',         'Cancelar ventas',                             'Ingresos'),
  ('autorizar_credito_excedido','Autorizar NC mayor al crédito del paciente',  'Ingresos'),
  ('identificar_pacientes',     'Identificar pacientes de cobros sin nombre',  'Ingresos'),
  ('editar_egresos',            'Editar partidas de egreso',                   'Egresos'),
  ('borrar_egresos',            'Eliminar partidas de egreso',                 'Egresos'),
  ('editar_productos',          'Editar productos y sus precios',              'Catálogo'),
  ('editar_pacientes',          'Registrar/editar pacientes',                  'Pacientes'),
  ('abrir_periodo_presupuesto', 'Abrir/cerrar periodos del presupuesto',       'Presupuesto'),
  ('sheets_access',             'Acceso directo / auditoría',                  'Sistema'),
  ('cf_capture',                'Capturar en Flujo de Efectivo',               'Flujo'),
  ('export_data',               'Exportar datos',                              'Sistema'),
  ('docs_enviar',               'Enviar documentos por correo',                'Documentos'),
  ('pagos_autorizar',           'Autorizar pagos',                             'Egresos'),
  ('ver_datos_sensibles',       'Ver nombres y datos sensibles',               'Privacidad'),
  ('cambiar_auto_refresh',      'Cambiar el auto-refresh',                     'Sistema')
on conflict (clave) do update set label = excluded.label, categoria = excluded.categoria;

-- ── Tenant Hestia ─────────────────────────────────────────────────────────
insert into core.tenants (slug, nombre, giro)
values ('hestia', 'Hestia Fertility', 'fertilidad')
on conflict (slug) do nothing;

-- ── Roles del tenant Hestia ───────────────────────────────────────────────
insert into core.roles (tenant_id, clave, nombre, es_super, solo_lectura)
select t.id, r.clave, r.nombre, r.es_super, r.solo_lectura
from core.tenants t
cross join (values
  ('admin',    'Administrador', true,  false),
  ('director', 'Dirección',     true,  false),
  ('gerente',  'Gerente',       false, false),
  ('operador', 'Operador',      false, false),
  ('captura',  'Captura',       false, false),
  ('socio',    'Socio',         false, true),
  ('viewer',   'Consulta',      false, true)
) as r(clave, nombre, es_super, solo_lectura)
where t.slug = 'hestia'
on conflict (tenant_id, clave) do nothing;

-- ── Permisos por rol (de ROLE_PERMS_DEFAULT). admin/director = '*'. ─────────
insert into core.rol_permiso (tenant_id, rol_clave, permiso_clave)
select t.id, m.rol_clave, m.permiso_clave
from core.tenants t
cross join (values
  ('admin','*'), ('director','*'),
  ('gerente','bank_create'),('gerente','bank_edit'),('gerente','bank_delete'),
  ('gerente','bank_capture'),('gerente','bank_liberado'),('gerente','sheets_access'),
  ('gerente','cf_capture'),('gerente','export_data'),('gerente','editar_ingresos'),
  ('gerente','editar_egresos'),('gerente','ver_datos_sensibles'),('gerente','identificar_pacientes'),('gerente','editar_pacientes'),
  ('operador','bank_capture'),('operador','bank_edit'),('operador','bank_liberado'),
  ('operador','cf_capture'),('operador','ver_datos_sensibles'),('operador','identificar_pacientes'),('operador','editar_pacientes'),
  ('captura','bank_capture'),('captura','cf_capture'),('captura','ver_datos_sensibles'),('captura','identificar_pacientes'),('captura','editar_pacientes'),
  ('socio','export_data')
  -- viewer: sin permisos
) as m(rol_clave, permiso_clave)
where t.slug = 'hestia'
on conflict do nothing;

-- ── PLANTILLA: enlazar un usuario staff (tras crearlo en Supabase Auth) ─────
-- 1) Crea el usuario en Supabase → Authentication → Add user (email+password).
-- 2) Copia su UUID y corre (descomentado):
-- insert into core.usuarios (id, tenant_id, email, nombre, alias, rol_clave)
-- select '<UUID-de-auth.users>', t.id, 'persona@hestiafertility.com', 'Nombre Completo', 'Alias', 'gerente'
-- from core.tenants t where t.slug = 'hestia';
