-- Alut: log of automatic status-change emails

CREATE TABLE IF NOT EXISTS public.alut_sent_emails (
  id SERIAL PRIMARY KEY,
  item_id INTEGER REFERENCES public.alut_order_items(id) ON DELETE SET NULL,
  order_id INTEGER REFERENCES public.alut_orders(id) ON DELETE SET NULL,

  trigger_status TEXT NOT NULL,
  recipient_type TEXT NOT NULL CHECK (recipient_type IN ('karin', 'client')),
  recipient_email TEXT NOT NULL,

  subject TEXT,
  success BOOLEAN DEFAULT FALSE,
  error_message TEXT,

  sent_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_alut_sent_emails_item ON public.alut_sent_emails(item_id);
CREATE INDEX IF NOT EXISTS idx_alut_sent_emails_order ON public.alut_sent_emails(order_id);

ALTER TABLE public.alut_sent_emails ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "alut_sent_emails_read_all" ON public.alut_sent_emails;
CREATE POLICY "alut_sent_emails_read_all" ON public.alut_sent_emails FOR SELECT USING (true);

DROP POLICY IF EXISTS "alut_sent_emails_write" ON public.alut_sent_emails;
CREATE POLICY "alut_sent_emails_write" ON public.alut_sent_emails FOR ALL USING (true) WITH CHECK (true);

GRANT SELECT ON public.alut_sent_emails TO anon;
GRANT ALL ON public.alut_sent_emails TO authenticated;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
