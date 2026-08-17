-- Rode este SQL completo no "SQL Editor" do Supabase (Run).
-- Cria a tabela de lembretes (cada um com data, texto e status de feito),
-- com as mesmas regras de segurança por usuário já usadas na tabela "entries".

create table if not exists lembretes (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) not null,
  data date not null,
  texto text not null,
  feito boolean not null default false,
  created_at timestamptz not null default now()
);

alter table lembretes enable row level security;

create policy "Usuarios veem so os proprios lembretes" on lembretes
  for select
  to authenticated
  using (auth.uid() = user_id);

create policy "Usuarios inserem so para si mesmos" on lembretes
  for insert
  to authenticated
  with check (auth.uid() = user_id);

create policy "Usuarios atualizam so os proprios lembretes" on lembretes
  for update
  to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

create policy "Usuarios excluem so os proprios lembretes" on lembretes
  for delete
  to authenticated
  using (auth.uid() = user_id);

-- Habilita atualização em tempo real (para lembretes aparecerem em outro
-- dispositivo sem precisar recarregar a página), igual já é feito em "entries".
alter publication supabase_realtime add table lembretes;
