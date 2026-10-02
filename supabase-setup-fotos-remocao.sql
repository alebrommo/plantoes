-- Coluna de fotos para remoções avulsas (tipo "remocao").
-- As remoções dentro de plantões HAPVIDA já guardam as fotos dentro da
-- coluna "remocoes" (jsonb) que já existe, sem precisar de mudança.
alter table entries add column if not exists fotos jsonb default '[]'::jsonb not null;

-- Bucket de armazenamento (privado — só o dono enxerga as próprias fotos).
insert into storage.buckets (id, name, public)
values ('remocao-fotos', 'remocao-fotos', false)
on conflict (id) do nothing;

create policy "Usuarios veem suas proprias fotos de remocao"
  on storage.objects for select
  to authenticated
  using (bucket_id = 'remocao-fotos' and (storage.foldername(name))[1] = auth.uid()::text);

create policy "Usuarios enviam suas proprias fotos de remocao"
  on storage.objects for insert
  to authenticated
  with check (bucket_id = 'remocao-fotos' and (storage.foldername(name))[1] = auth.uid()::text);

create policy "Usuarios excluem suas proprias fotos de remocao"
  on storage.objects for delete
  to authenticated
  using (bucket_id = 'remocao-fotos' and (storage.foldername(name))[1] = auth.uid()::text);
