-- Alut orders: transfer-to-shipments tracking
ALTER TABLE public.alut_orders
  ADD COLUMN IF NOT EXISTS transferred_to_shipments BOOLEAN DEFAULT FALSE,
  ADD COLUMN IF NOT EXISTS transferred_at TIMESTAMPTZ,
  ADD COLUMN IF NOT EXISTS shipment_task_id INTEGER REFERENCES public.tasks(id);

NOTIFY pgrst, 'reload schema';
