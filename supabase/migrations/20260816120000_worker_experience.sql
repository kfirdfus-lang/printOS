-- ============================================================
-- חבילה H שלב א׳ — צד העובד
-- ============================================================

alter table public.tasks
  add column if not exists department text,
  add column if not exists work_file_path text,
  add column if not exists work_file_name text,
  add column if not exists work_file_type text,
  add column if not exists worker_note text,
  add column if not exists completed_at timestamptz,
  add column if not exists completed_by text,
  add column if not exists issue_reported boolean not null default false,
  add column if not exists issue_text text,
  add column if not exists issue_reported_at timestamptz,
  add column if not exists first_viewed_at timestamptz;

create index if not exists idx_tasks_department
  on public.tasks (department);
create index if not exists idx_tasks_completed_at
  on public.tasks (completed_at);

-- ------------------------------------------------------------
-- מחלקות
-- ------------------------------------------------------------
create table if not exists public.departments (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  name text not null,
  color text not null default '#6B7280',
  sort_order integer not null default 0,
  has_rip boolean not null default true,
  is_active boolean not null default true,
  constraint departments_name_unique unique (name)
);

alter table public.departments enable row level security;
drop policy if exists "departments_all" on public.departments;
create policy "departments_all"
  on public.departments for all using (true) with check (true);

grant all on public.departments to anon, authenticated;

insert into public.departments (name, color, sort_order, has_rip) values
  ('הדפסה',   '#378ADD', 1, true),
  ('גימור',   '#1D9E75', 2, false),
  ('כריכה',   '#7F77DD', 3, false),
  ('משלוחים', '#BA7517', 4, false)
on conflict (name) do nothing;

-- ------------------------------------------------------------
-- קבצים בהמתנה לחיבור
-- ------------------------------------------------------------
create table if not exists public.pending_work_files (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  created_by_name text,
  bina_order_id text not null,
  file_path text not null,
  file_name text not null,
  file_type text,
  file_size_kb integer,
  worker_note text,
  due_date date,
  department text,
  linked_task_id uuid references public.tasks(id) on delete set null,
  linked_at timestamptz,
  status text not null default 'waiting'
    check (status in ('waiting','linked','cancelled'))
);

create index if not exists idx_pending_files_order
  on public.pending_work_files (bina_order_id);
create index if not exists idx_pending_files_status
  on public.pending_work_files (status);

alter table public.pending_work_files enable row level security;
drop policy if exists "pending_work_files_all" on public.pending_work_files;
create policy "pending_work_files_all"
  on public.pending_work_files for all using (true) with check (true);

grant all on public.pending_work_files to anon, authenticated;

notify pgrst, 'reload schema';
