-- ============================================================
-- מאגר אנשי קשר למשלוח לכל לקוח
-- ============================================================
CREATE TABLE IF NOT EXISTS public.client_delivery_contacts (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id uuid NOT NULL REFERENCES public.clients(id) ON DELETE CASCADE,
  contact_name text NOT NULL,
  contact_phone text NOT NULL,
  is_default boolean DEFAULT false,
  last_used_at timestamptz DEFAULT now(),
  usage_count integer DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  created_by text,
  UNIQUE(client_id, contact_phone)
);

CREATE INDEX IF NOT EXISTS idx_delivery_contacts_client 
  ON public.client_delivery_contacts(client_id);

CREATE INDEX IF NOT EXISTS idx_delivery_contacts_default 
  ON public.client_delivery_contacts(client_id, is_default)
  WHERE is_default = true;

CREATE INDEX IF NOT EXISTS idx_delivery_contacts_usage 
  ON public.client_delivery_contacts(client_id, usage_count DESC, last_used_at DESC);

ALTER TABLE public.client_delivery_contacts ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "client_delivery_contacts_all" ON public.client_delivery_contacts;
CREATE POLICY "client_delivery_contacts_all"
  ON public.client_delivery_contacts
  FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.client_delivery_contacts TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
