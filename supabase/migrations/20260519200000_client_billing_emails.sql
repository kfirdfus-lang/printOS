-- ============================================================
-- מאגר מיילים לגבייה/הנה"ח לכל לקוח
-- ============================================================
CREATE TABLE IF NOT EXISTS public.client_billing_emails (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  client_id uuid NOT NULL REFERENCES public.clients(id) ON DELETE CASCADE,
  email text NOT NULL,
  label text,
  is_default boolean DEFAULT false,
  last_used_at timestamptz DEFAULT now(),
  usage_count integer DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  created_by text,
  UNIQUE(client_id, email)
);

CREATE INDEX IF NOT EXISTS idx_billing_emails_client 
  ON public.client_billing_emails(client_id);

CREATE INDEX IF NOT EXISTS idx_billing_emails_default 
  ON public.client_billing_emails(client_id, is_default)
  WHERE is_default = true;

CREATE INDEX IF NOT EXISTS idx_billing_emails_usage 
  ON public.client_billing_emails(client_id, usage_count DESC, last_used_at DESC);

ALTER TABLE public.client_billing_emails ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "client_billing_emails_all" ON public.client_billing_emails;
CREATE POLICY "client_billing_emails_all"
  ON public.client_billing_emails
  FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.client_billing_emails TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
