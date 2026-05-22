-- ============================================================
-- מערכת בקשות גרפיות
-- ============================================================
CREATE TABLE IF NOT EXISTS public.design_requests (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),

  request_type text NOT NULL,

  title text NOT NULL,
  description text,

  status text DEFAULT 'pending',

  input_files jsonb DEFAULT '[]'::jsonb,
  output_files jsonb DEFAULT '[]'::jsonb,

  parameters jsonb DEFAULT '{}'::jsonb,

  created_at timestamptz DEFAULT now(),
  created_by text,
  started_at timestamptz,
  completed_at timestamptz,

  error_message text,
  processing_log jsonb DEFAULT '[]'::jsonb,

  client_id uuid,
  task_id uuid
);

CREATE INDEX IF NOT EXISTS idx_design_requests_status
  ON public.design_requests(status);
CREATE INDEX IF NOT EXISTS idx_design_requests_type
  ON public.design_requests(request_type);
CREATE INDEX IF NOT EXISTS idx_design_requests_created
  ON public.design_requests(created_at DESC);

ALTER TABLE public.design_requests ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "design_requests_all" ON public.design_requests;
CREATE POLICY "design_requests_all"
  ON public.design_requests FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.design_requests TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
