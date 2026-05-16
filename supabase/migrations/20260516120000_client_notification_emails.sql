-- Client notification email pool (per-client addresses with labels)

CREATE TABLE IF NOT EXISTS public.client_notification_emails (
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

CREATE INDEX IF NOT EXISTS idx_notif_emails_client
  ON public.client_notification_emails(client_id);
CREATE INDEX IF NOT EXISTS idx_notif_emails_default
  ON public.client_notification_emails(client_id, is_default)
  WHERE is_default = true;

ALTER TABLE public.client_notification_emails ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "client_notification_emails_all" ON public.client_notification_emails;
CREATE POLICY "client_notification_emails_all"
  ON public.client_notification_emails
  FOR ALL USING (true) WITH CHECK (true);
GRANT ALL ON public.client_notification_emails TO anon, authenticated;
