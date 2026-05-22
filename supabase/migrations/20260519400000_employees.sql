-- ============================================================
-- טבלת עובדים
-- ============================================================
CREATE TABLE IF NOT EXISTS public.employees (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),

  full_name text NOT NULL,
  phone text,
  birth_date date,

  is_active boolean DEFAULT true,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now(),

  notes text
);

CREATE INDEX IF NOT EXISTS idx_employees_active
  ON public.employees(is_active) WHERE is_active = true;
CREATE INDEX IF NOT EXISTS idx_employees_birth
  ON public.employees(birth_date) WHERE birth_date IS NOT NULL;

ALTER TABLE public.employees ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "employees_all" ON public.employees;
CREATE POLICY "employees_all"
  ON public.employees FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.employees TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
