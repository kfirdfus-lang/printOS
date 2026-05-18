-- ============================================================
-- תשתית קבצים משותפת
-- ============================================================

CREATE TABLE IF NOT EXISTS public.documents (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  category text,
  employee_id uuid,
  title text NOT NULL,
  description text,
  file_path text NOT NULL,
  file_name text NOT NULL,
  file_size bigint,
  file_type text,
  document_date date,
  expires_at date,
  uploaded_at timestamptz DEFAULT now(),
  uploaded_by text,
  tags text[],
  metadata jsonb
);

CREATE INDEX IF NOT EXISTS idx_documents_category 
  ON public.documents(category) WHERE category IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_documents_employee 
  ON public.documents(employee_id) WHERE employee_id IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_documents_expires 
  ON public.documents(expires_at) WHERE expires_at IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_documents_uploaded 
  ON public.documents(uploaded_at DESC);

ALTER TABLE public.documents ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "documents_all" ON public.documents;
CREATE POLICY "documents_all"
  ON public.documents FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.documents TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
