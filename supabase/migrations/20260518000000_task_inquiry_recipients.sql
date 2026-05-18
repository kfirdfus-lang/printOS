-- מאגר נמענים לפנייה לאחראי משימה
CREATE TABLE IF NOT EXISTS public.task_inquiry_recipients (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  email text UNIQUE NOT NULL,
  role text,
  is_default boolean DEFAULT false,
  last_used_at timestamptz DEFAULT now(),
  usage_count integer DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  created_by text
);

CREATE INDEX IF NOT EXISTS idx_inquiry_recipients_email
  ON public.task_inquiry_recipients(email);
CREATE INDEX IF NOT EXISTS idx_inquiry_recipients_usage
  ON public.task_inquiry_recipients(usage_count DESC, last_used_at DESC);

ALTER TABLE public.task_inquiry_recipients ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "task_inquiry_recipients_all" ON public.task_inquiry_recipients;
CREATE POLICY "task_inquiry_recipients_all"
  ON public.task_inquiry_recipients
  FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.task_inquiry_recipients TO anon, authenticated;
