-- ============================================================
-- שלב 1: כתובות משלוח של לקוחות
-- ============================================================
CREATE TABLE IF NOT EXISTS public.client_addresses (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id uuid NOT NULL REFERENCES public.clients(id) ON DELETE CASCADE,
  address text NOT NULL,
  label text,  -- "אסם", "בית", "סניף תל אביב" וכו'
  is_default boolean DEFAULT false,
  last_used_at timestamptz DEFAULT now(),
  usage_count integer DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  created_by text
);

CREATE INDEX IF NOT EXISTS idx_client_addresses_client ON public.client_addresses(client_id);
CREATE INDEX IF NOT EXISTS idx_client_addresses_default ON public.client_addresses(client_id, is_default) WHERE is_default = true;

ALTER TABLE public.client_addresses ENABLE ROW LEVEL SECURITY;
CREATE POLICY "client_addresses_all" ON public.client_addresses FOR ALL USING (true) WITH CHECK (true);
GRANT ALL ON public.client_addresses TO anon, authenticated;

-- ============================================================
-- שלב 2: עמודות חדשות לטבלת clients
-- ============================================================
ALTER TABLE public.clients 
  ADD COLUMN IF NOT EXISTS delivery_contact_name text,
  ADD COLUMN IF NOT EXISTS delivery_contact_phone text,
  ADD COLUMN IF NOT EXISTS delivery_email text,
  ADD COLUMN IF NOT EXISTS delivery_notes text;

-- ============================================================
-- שלב 3: סטטוס משלוח על משימות (לדשבורד משלוחים)
-- ============================================================
ALTER TABLE public.tasks
  ADD COLUMN IF NOT EXISTS delivery_status text DEFAULT 'pending' 
    CHECK (delivery_status IN ('pending', 'ready', 'in_transit', 'delivered', 'cancelled')),
  ADD COLUMN IF NOT EXISTS delivery_marked_at timestamptz,
  ADD COLUMN IF NOT EXISTS delivery_marked_by text;

CREATE INDEX IF NOT EXISTS idx_tasks_delivery_status ON public.tasks(delivery_status);
