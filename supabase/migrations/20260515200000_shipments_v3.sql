-- ============================================================
-- שלב 1: טבלת נהגים
-- ============================================================
CREATE TABLE IF NOT EXISTS public.drivers (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  phone text,
  color text DEFAULT '#0e7490',
  is_active boolean DEFAULT true,
  notes text,
  created_at timestamptz DEFAULT now()
);

ALTER TABLE public.drivers ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "drivers_all" ON public.drivers;
CREATE POLICY "drivers_all" ON public.drivers FOR ALL USING (true) WITH CHECK (true);
GRANT ALL ON public.drivers TO anon, authenticated;

INSERT INTO public.drivers (name, phone, color)
SELECT 'נהג 1', '', '#0e7490'
WHERE NOT EXISTS (SELECT 1 FROM public.drivers WHERE name = 'נהג 1');
INSERT INTO public.drivers (name, phone, color)
SELECT 'נהג 2', '', '#8b5cf6'
WHERE NOT EXISTS (SELECT 1 FROM public.drivers WHERE name = 'נהג 2');
INSERT INTO public.drivers (name, phone, color)
SELECT 'נהג 3', '', '#f59e0b'
WHERE NOT EXISTS (SELECT 1 FROM public.drivers WHERE name = 'נהג 3');

-- ============================================================
-- שלב 2: עמודות חדשות ל-tasks
-- ============================================================
ALTER TABLE public.tasks
  ADD COLUMN IF NOT EXISTS scheduled_delivery_date date,
  ADD COLUMN IF NOT EXISTS assigned_driver_id uuid REFERENCES public.drivers(id) ON DELETE SET NULL,
  ADD COLUMN IF NOT EXISTS delivery_order integer,
  ADD COLUMN IF NOT EXISTS delivery_type text DEFAULT 'local'
    CHECK (delivery_type IN ('local', 'courier', 'pickup')),
  ADD COLUMN IF NOT EXISTS delivery_address_snapshot text,
  ADD COLUMN IF NOT EXISTS delivery_contact_snapshot text,
  ADD COLUMN IF NOT EXISTS delivery_phone_snapshot text,
  ADD COLUMN IF NOT EXISTS delivery_special_notes text;

CREATE INDEX IF NOT EXISTS idx_tasks_scheduled_delivery ON public.tasks(scheduled_delivery_date)
  WHERE scheduled_delivery_date IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_tasks_assigned_driver ON public.tasks(assigned_driver_id)
  WHERE assigned_driver_id IS NOT NULL;
