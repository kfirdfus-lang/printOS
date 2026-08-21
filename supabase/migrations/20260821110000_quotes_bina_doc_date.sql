-- Quote document date from Bina (docDate) for history search / import
alter table public.quotes
  add column if not exists bina_doc_date date;

create index if not exists idx_quotes_bina_doc_date
  on public.quotes (bina_doc_date desc);

create unique index if not exists uq_quote_items_quote_line
  on public.quote_items (quote_id, line_number);

NOTIFY pgrst, 'reload schema';
