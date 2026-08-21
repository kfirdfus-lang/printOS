-- Package I: quotes archive separation for historical import
-- Apply via SQL editor / db query --linked. Do NOT use supabase db push.

alter table public.quotes
  add column if not exists is_archive boolean not null default false,
  add column if not exists archive_imported_at timestamptz;

create unique index if not exists uq_quotes_bina_doc_id
  on public.quotes (bina_doc_id)
  where bina_doc_id is not null;

create index if not exists idx_quotes_is_archive
  on public.quotes (is_archive);

create index if not exists idx_quotes_cust_name_lower
  on public.quotes (lower(bina_cust_name));

create index if not exists idx_quote_items_desc_lower
  on public.quote_items (lower(description));

NOTIFY pgrst, 'reload schema';
