-- ============================================================
-- חבילה K — אינדקסים לחישובי לקוח / תובנות
-- אין טבלאות חדשות
-- ============================================================

create index if not exists idx_tasks_client_date
  on public.tasks (client_name, bina_order_date desc)
  where bina_order_id is not null;

create index if not exists idx_tasks_bina_cust
  on public.tasks (bina_cust_id)
  where bina_cust_id is not null;

create index if not exists idx_task_items_dept_task
  on public.task_items (department, task_id);
