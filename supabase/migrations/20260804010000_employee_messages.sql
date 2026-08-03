-- Package G: history table for mass messages sent to employees.
-- Run manually in the Supabase SQL editor.

CREATE TABLE IF NOT EXISTS public.employee_messages (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  subject TEXT NOT NULL,
  body TEXT NOT NULL,
  recipient_type TEXT CHECK (recipient_type IN ('all', 'by_role', 'by_type', 'manual')),
  recipient_filter TEXT,
  recipient_employee_ids UUID[],
  recipient_emails TEXT[],
  sent_at TIMESTAMPTZ DEFAULT NOW(),
  sent_by_user_id UUID,
  sent_by_name TEXT,
  send_status TEXT DEFAULT 'sent' CHECK (send_status IN ('sent', 'partial', 'failed')),
  error_details TEXT
);

CREATE INDEX IF NOT EXISTS idx_employee_messages_sent_at
  ON public.employee_messages(sent_at DESC);

ALTER TABLE public.employee_messages ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "employee_messages_all" ON public.employee_messages;
CREATE POLICY "employee_messages_all" ON public.employee_messages FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.employee_messages TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
