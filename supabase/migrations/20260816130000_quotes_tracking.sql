-- ============================================================
-- חבילה H שלב ב׳ — מעקב הצעות (SQL בלבד; UI אחרי בדיקת שלב א׳)
-- NOTE: טבלת quotes כבר קיימת עם auto_status / closed_at וכו'.
--       כאן רק עמודות נוספות לזרימת package H v1.1 — בלי לשנות קיימות.
-- ============================================================

alter table public.quotes
  add column if not exists quote_status text not null default 'sent'
    check (quote_status in ('sent','approved','rejected','expired')),
  add column if not exists sent_at timestamptz default now(),
  add column if not exists status_changed_at timestamptz,
  add column if not exists converted_to_order_id text,
  add column if not exists converted_at timestamptz,
  add column if not exists follow_up_sent_at timestamptz;

create index if not exists idx_quotes_status on public.quotes (quote_status);
create index if not exists idx_quotes_sent_at on public.quotes (sent_at desc);

notify pgrst, 'reload schema';
