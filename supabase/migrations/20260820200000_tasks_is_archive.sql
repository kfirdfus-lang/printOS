-- Package I: historical archive separation + search indexes
-- Do NOT use supabase db push — apply via SQL editor or db query --linked

alter table public.tasks
  add column if not exists is_archive boolean not null default false,
  add column if not exists archive_imported_at timestamptz;

create index if not exists idx_tasks_is_archive
  on public.tasks (is_archive);

create index if not exists idx_tasks_client_name_lower
  on public.tasks (lower(client_name));

create index if not exists idx_tasks_bina_order_date
  on public.tasks (bina_order_date desc);

create index if not exists idx_task_items_desc_lower
  on public.task_items (lower(description));

-- Active lists: non-archive + not soft-archived
create index if not exists idx_tasks_active_created
  on public.tasks (created_at desc)
  where is_archive = false and archived_at is null;

NOTIFY pgrst, 'reload schema';
