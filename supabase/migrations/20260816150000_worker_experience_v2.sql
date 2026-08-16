-- ============================================================
-- חבילה H שלב א׳ v2 — מבוסס פריטים
-- ============================================================

-- ניקוי גרסה קודמת
drop table if exists public.departments cascade;
alter table public.tasks drop column if exists department;
alter table public.pending_work_files drop column if exists department;

-- ------------------------------------------------------------
-- הרחבת task_items — מעקב עובד
-- ------------------------------------------------------------
alter table public.task_items
  add column if not exists completed_by text,
  add column if not exists first_viewed_at timestamptz,
  add column if not exists issue_reported boolean not null default false,
  add column if not exists issue_text text,
  add column if not exists issue_reported_at timestamptz;

create index if not exists idx_task_items_department
  on public.task_items (department);
create index if not exists idx_task_items_status
  on public.task_items (status);
create index if not exists idx_task_items_task
  on public.task_items (task_id);

-- ------------------------------------------------------------
-- הגדרות מחלקה (רק תצוגה — לא רשימת המחלקות עצמה)
-- ------------------------------------------------------------
create table if not exists public.department_settings (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  name text not null,
  color text not null default '#6B7280',
  has_rip boolean not null default true,
  sort_order integer not null default 0,
  constraint department_settings_name_unique unique (name)
);

alter table public.department_settings enable row level security;
drop policy if exists "department_settings_all" on public.department_settings;
create policy "department_settings_all"
  on public.department_settings for all using (true) with check (true);

grant all on public.department_settings to anon, authenticated;

insert into public.department_settings (name, color, has_rip, sort_order) values
  ('פורמט רחב',          '#378ADD', true,  1),
  ('דיגיטלי צבעוני',     '#1D9E75', true,  2),
  ('דיגיטלי שחור לבן',   '#6B7280', true,  3),
  ('אופסט',              '#7F77DD', true,  4),
  ('ביגוד ומוצרי פרסום', '#BA7517', true,  5),
  ('עבודות חוץ',         '#C2410C', false, 6),
  ('משלוחים',            '#0891B2', false, 7)
on conflict (name) do nothing;

-- ------------------------------------------------------------
-- קבצים בהמתנה (נשאר מגרסה קודמת אם כבר קיים)
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

-- שדות קובץ על tasks (אם עדיין לא קיימים מגרסה קודמת)
alter table public.tasks
  add column if not exists work_file_path text,
  add column if not exists work_file_name text,
  add column if not exists work_file_type text,
  add column if not exists worker_note text,
  add column if not exists completed_by text,
  add column if not exists first_viewed_at timestamptz;

notify pgrst, 'reload schema';
