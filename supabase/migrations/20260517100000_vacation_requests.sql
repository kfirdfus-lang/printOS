-- ============================================================
-- טפסי בקשת חופש דיגיטליים
-- ============================================================
CREATE TABLE IF NOT EXISTS public.vacation_requests (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),

  employee_name text NOT NULL,
  employee_id_last4 text NOT NULL,
  employee_full_id text,

  start_date date NOT NULL,
  end_date date NOT NULL,
  total_days integer,

  reason_type text NOT NULL
    CHECK (reason_type IN ('vacation', 'sick', 'reserve', 'other')),
  reason_notes text,

  signature_data text NOT NULL,
  signed_at timestamptz DEFAULT now(),

  status text DEFAULT 'pending'
    CHECK (status IN ('pending', 'approved', 'rejected')),
  reviewed_at timestamptz,
  reviewed_by text,
  reviewer_notes text,

  submitted_at timestamptz DEFAULT now(),
  ip_address text,
  user_agent text
);

CREATE INDEX IF NOT EXISTS idx_vacation_employee
  ON public.vacation_requests(employee_name, employee_id_last4);
CREATE INDEX IF NOT EXISTS idx_vacation_status
  ON public.vacation_requests(status, submitted_at DESC);
CREATE INDEX IF NOT EXISTS idx_vacation_dates
  ON public.vacation_requests(start_date, end_date);

CREATE TABLE IF NOT EXISTS public.vacation_form_config (
  id integer PRIMARY KEY DEFAULT 1,
  form_slug text UNIQUE NOT NULL,
  is_active boolean DEFAULT true,
  created_at timestamptz DEFAULT now(),
  CONSTRAINT single_config CHECK (id = 1)
);

INSERT INTO public.vacation_form_config (id, form_slug, is_active)
VALUES (1, 'natalie-vacation', true)
ON CONFLICT (id) DO NOTHING;

ALTER TABLE public.vacation_requests ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.vacation_form_config ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "vacation_requests_all" ON public.vacation_requests;
CREATE POLICY "vacation_requests_all"
  ON public.vacation_requests
  FOR ALL USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "vacation_config_all" ON public.vacation_form_config;
CREATE POLICY "vacation_config_all"
  ON public.vacation_form_config
  FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.vacation_requests TO anon, authenticated;
GRANT ALL ON public.vacation_form_config TO anon, authenticated;
