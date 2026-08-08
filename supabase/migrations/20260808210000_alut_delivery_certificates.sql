-- Delivery certificates for Alut items (digital customer signature on delivery/pickup)
-- NOTE: alut_orders.id and alut_order_items.id are INTEGER (SERIAL), not UUID.

CREATE TABLE IF NOT EXISTS public.alut_delivery_certificates (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  order_id INTEGER NOT NULL REFERENCES public.alut_orders(id) ON DELETE CASCADE,
  item_id INTEGER REFERENCES public.alut_order_items(id) ON DELETE CASCADE,

  recipient_name TEXT NOT NULL,
  recipient_id_number TEXT,

  signature_data TEXT NOT NULL,

  photo_url TEXT,

  signed_at TIMESTAMPTZ DEFAULT NOW(),
  signed_by_user_id TEXT,
  signed_by_user_name TEXT,

  notes TEXT,

  pdf_url TEXT,

  created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_alut_delivery_certs_order
  ON public.alut_delivery_certificates(order_id);
CREATE INDEX IF NOT EXISTS idx_alut_delivery_certs_item
  ON public.alut_delivery_certificates(item_id);
CREATE INDEX IF NOT EXISTS idx_alut_delivery_certs_date
  ON public.alut_delivery_certificates(signed_at DESC);

ALTER TABLE public.alut_delivery_certificates ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "alut_delivery_certs_all" ON public.alut_delivery_certificates;
CREATE POLICY "alut_delivery_certs_all" ON public.alut_delivery_certificates
  FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.alut_delivery_certificates TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
