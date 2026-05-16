-- Client notification emails on tasks (ready / shipped / delivered)

ALTER TABLE public.tasks
  ADD COLUMN IF NOT EXISTS notification_email text,
  ADD COLUMN IF NOT EXISTS notif_ready_sent_at timestamptz,
  ADD COLUMN IF NOT EXISTS notif_shipped_sent_at timestamptz,
  ADD COLUMN IF NOT EXISTS notif_delivered_sent_at timestamptz;

CREATE TABLE IF NOT EXISTS public.client_notifications_log (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  task_id uuid REFERENCES public.tasks(id) ON DELETE SET NULL,
  notification_type text NOT NULL
    CHECK (notification_type IN ('ready', 'shipped', 'delivered')),
  recipient_email text NOT NULL,
  cc_emails text[],
  subject text,
  body text,
  sent_at timestamptz DEFAULT now(),
  sent_by text,
  status text DEFAULT 'sent' CHECK (status IN ('sent', 'failed')),
  error_message text
);

CREATE INDEX IF NOT EXISTS idx_notifications_task ON public.client_notifications_log(task_id);
CREATE INDEX IF NOT EXISTS idx_notifications_sent_at ON public.client_notifications_log(sent_at DESC);

ALTER TABLE public.client_notifications_log ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "client_notifications_all" ON public.client_notifications_log;
CREATE POLICY "client_notifications_all" ON public.client_notifications_log
  FOR ALL USING (true) WITH CHECK (true);
GRANT ALL ON public.client_notifications_log TO anon, authenticated;
