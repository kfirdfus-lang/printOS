-- G3: snooze emails until a date (PrintOS-side; Gmail archive + restore).
-- Run in SQL editor (do not db push).

create table if not exists public.gmail_snoozed (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  message_id text not null,
  thread_id text,
  snooze_until timestamptz not null,
  note text,
  released boolean not null default false,
  constraint gmail_snoozed_msg_unique unique (message_id)
);

create index if not exists idx_gmail_snoozed_until
  on public.gmail_snoozed (snooze_until) where released = false;

alter table public.gmail_snoozed enable row level security;
revoke all on public.gmail_snoozed from anon, authenticated;

comment on table public.gmail_snoozed is
  'Snoozed Gmail messages restored to INBOX by daily cron. Service-role only.';

NOTIFY pgrst, 'reload schema';
