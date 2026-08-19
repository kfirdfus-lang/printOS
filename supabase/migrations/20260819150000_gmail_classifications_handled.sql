-- G3: track handled vs pending Gmail classifications.
-- Run in SQL editor (do not db push).

alter table public.gmail_classifications
  add column if not exists handled boolean not null default false,
  add column if not exists handled_at timestamptz,
  add column if not exists handled_by text,
  add column if not exists handled_reason text,
  add column if not exists created_order_id text;

create index if not exists idx_gmail_class_handled
  on public.gmail_classifications (handled);

NOTIFY pgrst, 'reload schema';
