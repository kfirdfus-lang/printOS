-- Store Bina invoice number (invNumber) on synced orders.
alter table public.tasks
  add column if not exists bina_invoice_number integer;

create index if not exists idx_tasks_invoice
  on public.tasks (bina_invoice_number)
  where bina_invoice_number is not null;
