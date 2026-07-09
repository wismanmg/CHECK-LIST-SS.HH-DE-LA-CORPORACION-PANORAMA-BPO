-- =====================================================================
-- CHECK LIST SS.HH DE LA CORPORACION - Tablas para Supabase
-- Pegar y ejecutar en: Supabase -> SQL Editor -> New query -> Run
-- =====================================================================

create table if not exists observaciones (
  id             bigint generated always as identity primary key,
  marca_temporal timestamptz not null default now(),
  edificio       text not null default '',
  piso           text not null default '',
  ubicacion      text not null default '',
  empresas       text not null default '',
  sshh           text not null default '',
  lavatorio      text not null default '',
  mesa           text not null default '',
  inodoro        text not null default '',
  puertas        text not null default '',
  urinario       text not null default '',
  descripcion    text not null default '',
  comentario     text not null default '',
  dispensador    text not null default '',
  comentario_fmi text not null default '',
  proveedor      text not null default '',
  estatus        text not null default 'PENDIENTE',
  po             text not null default '',
  monto          numeric,
  prioridad      text not null default ''
);

create table if not exists historial (
  id             bigint generated always as identity primary key,
  fecha          timestamptz not null default now(),
  caso_id        bigint,
  usuario        text not null default '',
  campo          text not null default '',
  valor_anterior text not null default '',
  valor_nuevo    text not null default ''
);

-- La app se conecta con la clave service_role (servidor), que no depende de RLS.
-- Activamos RLS sin políticas para que nadie pueda leer/escribir con la clave pública (anon).
alter table observaciones enable row level security;
alter table historial enable row level security;
