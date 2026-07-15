-- Rode este SQL no "SQL Editor" do seu projeto Supabase (supabase.com).

create table if not exists entries (
  id text primary key,
  day_key text not null,
  type text not null,
  value numeric not null,
  color text,
  pago boolean not null default false,
  local text,
  i_h text,
  i_m text,
  f_h text,
  f_m text,
  turno text,
  empresa text,
  paciente text,
  origem text,
  destino text,
  obs text,
  created_at timestamptz not null default now()
);

create index if not exists entries_day_key_idx on entries (day_key);

-- Libera leitura e escrita para quem tem só a chave "anon" (sem login),
-- já que o app foi combinado para funcionar sem autenticação.
alter table entries enable row level security;

create policy "Acesso anônimo total" on entries
  for all
  to anon
  using (true)
  with check (true);

-- Habilita realtime (atualização automática entre dispositivos) nesta tabela.
alter publication supabase_realtime add table entries;
