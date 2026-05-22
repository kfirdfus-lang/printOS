ALTER TABLE public.tasks
  ADD COLUMN IF NOT EXISTS last_bina_sync timestamptz;

CREATE INDEX IF NOT EXISTS idx_tasks_last_bina_sync
  ON public.tasks(last_bina_sync)
  WHERE bina_order_id IS NOT NULL;

NOTIFY pgrst, 'reload schema';
