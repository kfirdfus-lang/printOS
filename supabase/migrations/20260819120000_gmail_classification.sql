-- G2: Gmail AI classifications. Service-role only.
-- Run in the SQL editor (do not db push).

create table if not exists public.gmail_classifications (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),

  message_id text not null,
  thread_id text,

  category text not null
    check (category in ('order','quote_request','supplier_invoice',
                        'design_approval','general','irrelevant')),
  confidence numeric(3,2),
  reason text,

  client_name text,
  extracted_data jsonb,

  model_used text,
  classified_at timestamptz not null default now(),

  user_corrected_category text,
  corrected_at timestamptz,

  constraint gmail_classifications_msg_unique unique (message_id)
);

create index if not exists idx_gmail_class_category
  on public.gmail_classifications (category);
create index if not exists idx_gmail_class_created
  on public.gmail_classifications (created_at desc);

alter table public.gmail_classifications enable row level security;
revoke all on public.gmail_classifications from anon, authenticated;

comment on table public.gmail_classifications is
  'Gmail AI classifications for the read-only inbox. Service-role only.';

NOTIFY pgrst, 'reload schema';
