-- Alut: at_itzik status + itzik work order audit log

ALTER TABLE public.alut_order_items
  DROP CONSTRAINT IF EXISTS alut_order_items_status_check;

ALTER TABLE public.alut_order_items
  ADD CONSTRAINT alut_order_items_status_check
  CHECK (status IN (
    'new',
    'design_in_progress',
    'design_sent',
    'design_approved',
    'in_print',
    'at_davach',
    'at_itzik',
    'ready_to_ship',
    'shipped',
    'delivered',
    'cancelled'
  ));

ALTER TABLE public.alut_order_items
  ADD COLUMN IF NOT EXISTS itzik_sent_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS itzik_received_at TIMESTAMPTZ;

CREATE TABLE IF NOT EXISTS public.alut_itzik_orders (
  id SERIAL PRIMARY KEY,
  item_id INTEGER NOT NULL REFERENCES public.alut_order_items(id) ON DELETE CASCADE,
  order_id INTEGER NOT NULL REFERENCES public.alut_orders(id) ON DELETE CASCADE,
  calendar_type TEXT NOT NULL,
  quantity INTEGER NOT NULL,
  cartons_count INTEGER NOT NULL,
  pdf_filename TEXT,
  notes TEXT,
  generated_at TIMESTAMPTZ DEFAULT NOW(),
  generated_by TEXT DEFAULT 'system'
);

CREATE INDEX IF NOT EXISTS idx_itzik_orders_item ON public.alut_itzik_orders(item_id);
CREATE INDEX IF NOT EXISTS idx_itzik_orders_order ON public.alut_itzik_orders(order_id);

ALTER TABLE public.alut_itzik_orders ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "alut_itzik_read_all" ON public.alut_itzik_orders;
CREATE POLICY "alut_itzik_read_all" ON public.alut_itzik_orders FOR SELECT USING (true);

DROP POLICY IF EXISTS "alut_itzik_write" ON public.alut_itzik_orders;
CREATE POLICY "alut_itzik_write" ON public.alut_itzik_orders FOR ALL USING (true) WITH CHECK (true);

GRANT SELECT ON public.alut_itzik_orders TO anon;
GRANT ALL ON public.alut_itzik_orders TO authenticated;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
