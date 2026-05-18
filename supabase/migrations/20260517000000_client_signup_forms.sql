-- ============================================================
-- טופס פתיחת לקוח דיגיטלי
-- ============================================================
CREATE TABLE IF NOT EXISTS public.client_signup_forms (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  token text UNIQUE NOT NULL,

  client_type text NOT NULL CHECK (client_type IN ('business', 'private')),

  business_name text,
  business_id text,
  full_name text,
  personal_id text,

  phone text NOT NULL,
  email text NOT NULL,
  contact_name text,

  billing_address text NOT NULL,
  shipping_address text,

  orders_email text,
  payment_terms text,
  notes text,

  status text DEFAULT 'pending'
    CHECK (status IN ('pending', 'approved', 'rejected', 'needs_review')),
  submitted_at timestamptz,
  reviewed_at timestamptz,
  reviewed_by text,
  reviewer_notes text,
  bina_customer_id text,
  bina_synced_at timestamptz,

  sent_by text,
  sent_to_name text,
  sent_to_contact text,

  created_at timestamptz DEFAULT now(),
  expires_at timestamptz DEFAULT now() + interval '30 days'
);

CREATE INDEX IF NOT EXISTS idx_signup_forms_token
  ON public.client_signup_forms(token);
CREATE INDEX IF NOT EXISTS idx_signup_forms_status
  ON public.client_signup_forms(status, submitted_at DESC);

ALTER TABLE public.client_signup_forms ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "signup_forms_all" ON public.client_signup_forms;
CREATE POLICY "signup_forms_all"
  ON public.client_signup_forms
  FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.client_signup_forms TO anon, authenticated;
