alter table public.quotes add column if not exists contact_person text;
alter table public.quotes add column if not exists sales_agent text;

comment on column public.quotes.contact_person is 'Lead / client-side contact name (מודאל הצעות).';
comment on column public.quotes.sales_agent is 'Internal sales rep name (מודאל הצעות).';
