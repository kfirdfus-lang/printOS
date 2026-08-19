-- G2 part 5 (later): sent-mail log + user signature.
-- Run after gmail.send scope is added. Do not db push.

create table if not exists public.gmail_sent_log (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),

  sent_by text,
  to_addresses text[],
  cc_addresses text[],
  subject text,
  body_preview text,

  in_reply_to_message_id text,
  thread_id text,
  gmail_message_id text,

  attachment_names text[],
  printos_document_type text,
  printos_document_id text,

  status text not null default 'sent'
    check (status in ('sent','failed')),
  error_text text
);

create index if not exists idx_gmail_sent_created
  on public.gmail_sent_log (created_at desc);
create index if not exists idx_gmail_sent_thread
  on public.gmail_sent_log (thread_id);

alter table public.gmail_sent_log enable row level security;
revoke all on public.gmail_sent_log from anon, authenticated;

alter table public.users
  add column if not exists email_signature text;

NOTIFY pgrst, 'reload schema';
