-- Package H2: TV dashboard gallery + monthly target settings

CREATE TABLE IF NOT EXISTS public.tv_gallery_projects (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  title TEXT NOT NULL,
  client_name TEXT,
  work_type TEXT,
  image_url TEXT NOT NULL,
  display_order INT DEFAULT 100,
  is_active BOOLEAN DEFAULT true,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_tv_gallery_active
  ON public.tv_gallery_projects(is_active, display_order);

ALTER TABLE public.tv_gallery_projects ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "tv_gallery_read_all" ON public.tv_gallery_projects;
CREATE POLICY "tv_gallery_read_all" ON public.tv_gallery_projects
  FOR SELECT USING (true);

DROP POLICY IF EXISTS "tv_gallery_write" ON public.tv_gallery_projects;
CREATE POLICY "tv_gallery_write" ON public.tv_gallery_projects
  FOR ALL USING (true) WITH CHECK (true);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.tv_gallery_projects TO anon, authenticated;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables
    WHERE pubname = 'supabase_realtime'
      AND schemaname = 'public'
      AND tablename = 'tv_gallery_projects'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE public.tv_gallery_projects;
  END IF;
END $$;

CREATE TABLE IF NOT EXISTS public.tv_dashboard_settings (
  id INT PRIMARY KEY DEFAULT 1,
  monthly_target NUMERIC DEFAULT 400000,
  updated_at TIMESTAMPTZ DEFAULT NOW(),
  CHECK (id = 1)
);

INSERT INTO public.tv_dashboard_settings (id, monthly_target)
VALUES (1, 400000)
ON CONFLICT (id) DO NOTHING;

ALTER TABLE public.tv_dashboard_settings ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "tv_settings_all" ON public.tv_dashboard_settings;
CREATE POLICY "tv_settings_all" ON public.tv_dashboard_settings
  FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.tv_dashboard_settings TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
