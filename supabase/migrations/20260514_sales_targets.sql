CREATE TABLE IF NOT EXISTS sales_targets (
  id SERIAL PRIMARY KEY,
  year INTEGER NOT NULL UNIQUE,
  annual_target NUMERIC(12, 2) NOT NULL,
  opening_amount NUMERIC(12, 2) NOT NULL,
  opening_date DATE NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

INSERT INTO sales_targets (year, annual_target, opening_amount, opening_date)
VALUES (2026, 5800000, 1328024, '2026-04-30')
ON CONFLICT (year) DO NOTHING;

ALTER TABLE public.sales_targets ENABLE ROW LEVEL SECURITY;

CREATE POLICY "sales_targets_select_all"
ON public.sales_targets
FOR SELECT
USING (true);

GRANT SELECT ON public.sales_targets TO anon, authenticated;
