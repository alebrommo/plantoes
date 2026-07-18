-- Rode este SQL completo no "SQL Editor" do Supabase (Run).
-- Ele: adiciona a coluna de dono, troca o acesso aberto por acesso
-- só de quem estiver logado, e migra os dados que já existem para
-- a sua conta (alebrommo@hotmail.com).

-- 1. Adiciona a coluna que identifica o dono de cada registro
alter table entries add column if not exists user_id uuid references auth.users(id);

-- 2. Remove a política antiga (acesso anônimo total)
drop policy if exists "Acesso anônimo total" on entries;

-- 3. Cria as 4 políticas novas: cada usuário logado só vê/edita o que é dele
create policy "Usuarios veem so os proprios registros" on entries
  for select
  to authenticated
  using (auth.uid() = user_id);

create policy "Usuarios inserem so para si mesmos" on entries
  for insert
  to authenticated
  with check (auth.uid() = user_id);

create policy "Usuarios atualizam so os proprios registros" on entries
  for update
  to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

create policy "Usuarios excluem so os proprios registros" on entries
  for delete
  to authenticated
  using (auth.uid() = user_id);

-- 4. Migra os registros que já existiam (sem dono) para a sua conta
update entries set user_id = '0f7aa9f5-82c9-4d30-adc8-21d7cc437cf9' where user_id is null;
