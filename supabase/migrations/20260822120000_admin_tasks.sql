-- ============================================================
-- חבילה J — משימות מנהלה
-- ============================================================

create table if not exists public.admin_tasks (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  created_by text,

  title text not null,
  details text,

  status text not null default 'open'
    check (status in ('open','in_progress','done')),

  assignee text not null default 'both'
    check (assignee in ('kfir','natalie','both')),

  due_date date,

  -- חזרתיות
  is_recurring boolean not null default false,
  recur_type text
    check (recur_type in ('monthly','yearly','custom_days')),
  recur_day integer,          -- יום בחודש (1-31) עבור monthly
  recur_month integer,        -- חודש (1-12) עבור yearly
  recur_days integer,         -- מספר ימים עבור custom_days
  parent_task_id uuid references public.admin_tasks(id) on delete set null,

  completed_at timestamptz,
  completed_by text,

  updated_at timestamptz not null default now()
);

create index if not exists idx_admin_tasks_status
  on public.admin_tasks (status);
create index if not exists idx_admin_tasks_due
  on public.admin_tasks (due_date);
create index if not exists idx_admin_tasks_assignee
  on public.admin_tasks (assignee);

alter table public.admin_tasks enable row level security;
drop policy if exists "admin_tasks_all" on public.admin_tasks;
create policy "admin_tasks_all"
  on public.admin_tasks for all using (true) with check (true);


-- ------------------------------------------------------------
-- הערות — רצף עם חותמת זמן, לא שדה שנדרס
-- ------------------------------------------------------------
create table if not exists public.admin_task_notes (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  task_id uuid not null references public.admin_tasks(id) on delete cascade,
  author text,
  note text not null
);

create index if not exists idx_admin_notes_task
  on public.admin_task_notes (task_id, created_at desc);

alter table public.admin_task_notes enable row level security;
drop policy if exists "admin_task_notes_all" on public.admin_task_notes;
create policy "admin_task_notes_all"
  on public.admin_task_notes for all using (true) with check (true);
