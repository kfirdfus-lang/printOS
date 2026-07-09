-- Alut: davach work order PDF audit log

CREATE TABLE IF NOT EXISTS public.alut_davach_orders (
  id SERIAL PRIMARY KEY,
  item_id INTEGER NOT NULL REFERENCES public.alut_order_items(id) ON DELETE CASCADE,
  order_id INTEGER NOT NULL REFERENCES public.alut_orders(id) ON DELETE CASCADE,

  calendar_type TEXT NOT NULL,
  quantity INTEGER NOT NULL,
  cartons_count INTEGER NOT NULL,
  cost_per_unit NUMERIC(10,2) NOT NULL,
  total_cost NUMERIC(10,2) NOT NULL,

  pdf_url TEXT,
  pdf_filename TEXT,

  notes TEXT,

  generated_at TIMESTAMPTZ DEFAULT NOW(),
  generated_by TEXT DEFAULT 'system'
);

CREATE INDEX IF NOT EXISTS idx_davach_orders_item ON public.alut_davach_orders(item_id);
CREATE INDEX IF NOT EXISTS idx_davach_orders_order ON public.alut_davach_orders(order_id);

ALTER TABLE public.alut_davach_orders ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "alut_davach_read_all" ON public.alut_davach_orders;
CREATE POLICY "alut_davach_read_all" ON public.alut_davach_orders FOR SELECT USING (true);

DROP POLICY IF EXISTS "alut_davach_write" ON public.alut_davach_orders;
CREATE POLICY "alut_davach_write" ON public.alut_davach_orders FOR ALL USING (true) WITH CHECK (true);

GRANT SELECT ON public.alut_davach_orders TO anon;
GRANT ALL ON public.alut_davach_orders TO authenticated;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
