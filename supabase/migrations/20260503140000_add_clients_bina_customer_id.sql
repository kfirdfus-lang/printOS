-- clients.notes already exists; add Bina customer identifier for PrintOS sync.
alter table public.clients
  add column if not exists bina_customer_id text;

comment on column public.clients.bina_customer_id is 'Identifier returned by Bina when creating a customer (e.g. kuponId / requestId).';

create unique index if not exists clients_bina_customer_id_uidx
  on public.clients (bina_customer_id)
  where bina_customer_id is not null and length(trim(bina_customer_id)) > 0;
