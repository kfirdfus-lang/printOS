-- G3: learned email → client mappings. Service-role only.
-- Run in SQL editor (do not db push).

create table if not exists public.client_email_aliases (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  client_id uuid not null references public.clients(id) on delete cascade,
  email text not null,
  constraint client_email_aliases_unique unique (email)
);

create index if not exists idx_client_aliases_email
  on public.client_email_aliases (email);

alter table public.client_email_aliases enable row level security;
revoke all on public.client_email_aliases from anon, authenticated;

comment on table public.client_email_aliases is
  'Gmail sender email aliases learned from manual client picks. Service-role only.';

NOTIFY pgrst, 'reload schema';
