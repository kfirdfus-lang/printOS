-- פריטים לפי שורה בהזמנה בינה — מחלקה לפי קוד ב-itemId (1–7)

CREATE TABLE IF NOT EXISTS public.task_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  task_id uuid NOT NULL REFERENCES public.tasks(id) ON DELETE CASCADE,
  bina_order_id bigint NOT NULL,
  line_number int NOT NULL,
  bina_item_code text,
  department text,
  description text NOT NULL,
  quantity numeric(10, 2) DEFAULT 0,
  price numeric(12, 2) DEFAULT 0,
  total numeric(12, 2) DEFAULT 0,
  status text NOT NULL DEFAULT 'בעבודה',
  notes text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  completed_at timestamptz,
  UNIQUE (bina_order_id, line_number)
);

CREATE INDEX IF NOT EXISTS idx_task_items_task_id ON public.task_items(task_id);
CREATE INDEX IF NOT EXISTS idx_task_items_department ON public.task_items(department);
CREATE INDEX IF NOT EXISTS idx_task_items_status ON public.task_items(status);

ALTER TABLE public.tasks
  ADD COLUMN IF NOT EXISTS has_items boolean NOT NULL DEFAULT false;

COMMENT ON TABLE public.task_items IS 'פריטים בהזמנה — כל פריט עם מחלקה משלה (קוד בינה ב-itemId)';
COMMENT ON COLUMN public.task_items.bina_item_code IS 'קוד מחלקה כפי שנרשם בבינה (1–7)';

CREATE OR REPLACE FUNCTION public.task_items_set_updated_at()
RETURNS trigger
LANGUAGE plpgsql
AS $$
BEGIN
  NEW.updated_at := now();
  RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS task_items_set_updated_at ON public.task_items;
CREATE TRIGGER task_items_set_updated_at
  BEFORE UPDATE ON public.task_items
  FOR EACH ROW
  execute function public.task_items_set_updated_at();
