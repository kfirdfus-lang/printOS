-- תזמון יומי למשימות (נפרד מ-due_date)
ALTER TABLE public.tasks
  ADD COLUMN IF NOT EXISTS scheduled_for_date date,
  ADD COLUMN IF NOT EXISTS scheduled_order integer DEFAULT 0;

CREATE INDEX IF NOT EXISTS idx_tasks_scheduled_for
  ON public.tasks(scheduled_for_date)
  WHERE scheduled_for_date IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_tasks_scheduled_order
  ON public.tasks(scheduled_for_date, scheduled_order);

NOTIFY pgrst, 'reload schema';
