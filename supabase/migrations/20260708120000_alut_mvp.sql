-- ============================================================
-- אלוט MVP: מחירים, פרויקטים, הזמנות, היסטוריית סטטוסים
-- ============================================================

CREATE TABLE IF NOT EXISTS public.alut_pricing (
  id SERIAL PRIMARY KEY,
  key TEXT UNIQUE NOT NULL,
  label TEXT NOT NULL,
  value NUMERIC(10, 2) NOT NULL,
  unit TEXT DEFAULT 'per_unit',
  category TEXT NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

INSERT INTO public.alut_pricing (key, label, value, unit, category) VALUES
  ('price_hard', 'לוח שולחני קשיח (מחיר ליחידה)', 7.50, 'per_unit', 'client_prices'),
  ('price_duplex', 'לוח שולחני דופלקס (מחיר ליחידה)', 4.50, 'per_unit', 'client_prices'),
  ('price_wall', 'לוח קיר (מחיר ליחידה)', 9.00, 'per_unit', 'client_prices'),
  ('price_design', 'הכנת סקיצה - השתלת לוגו וברכה', 70.00, 'per_order', 'client_prices'),
  ('price_shipping_carton', 'משלוח לקרטון', 40.00, 'per_carton', 'client_prices'),
  ('cost_davach_per_unit', 'עלות דבח (ליחידה - דופלקס וקיר)', 1.90, 'per_unit', 'davach_costs'),
  ('carton_hard', 'יחידות בקרטון - קשיח', 70, 'units', 'carton_config'),
  ('carton_duplex', 'יחידות בקרטון - דופלקס', 100, 'units', 'carton_config'),
  ('carton_wall', 'יחידות בקרטון - קיר', 100, 'units', 'carton_config')
ON CONFLICT (key) DO NOTHING;

CREATE TABLE IF NOT EXISTS public.alut_projects (
  id SERIAL PRIMARY KEY,
  year INTEGER UNIQUE NOT NULL,
  name TEXT NOT NULL,
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

INSERT INTO public.alut_projects (year, name, is_active) VALUES
  (2026, 'אלוט 2026', TRUE)
ON CONFLICT (year) DO NOTHING;

CREATE TABLE IF NOT EXISTS public.alut_orders (
  id SERIAL PRIMARY KEY,
  project_id INTEGER REFERENCES public.alut_projects(id) DEFAULT 1,
  order_number INTEGER NOT NULL,

  company_name TEXT NOT NULL,
  contact_name TEXT,
  contact_phone TEXT,
  contact_email TEXT,
  delivery_address TEXT,
  is_pickup BOOLEAN DEFAULT FALSE,

  calendar_type TEXT NOT NULL CHECK (calendar_type IN ('hard', 'duplex', 'wall')),
  quantity INTEGER NOT NULL,
  needs_design BOOLEAN DEFAULT TRUE,

  unit_price NUMERIC(10,2) NOT NULL,
  design_price NUMERIC(10,2) DEFAULT 0,
  shipping_price NUMERIC(10,2) DEFAULT 0,
  cartons_count INTEGER DEFAULT 0,
  total_price NUMERIC(10,2) DEFAULT 0,

  davach_cost NUMERIC(10,2) DEFAULT 0,
  davach_sent_at TIMESTAMPTZ,
  davach_received_at TIMESTAMPTZ,

  status TEXT DEFAULT 'new' CHECK (status IN (
    'new',
    'design_in_progress',
    'design_sent',
    'design_approved',
    'in_print',
    'at_davach',
    'ready_to_ship',
    'shipped',
    'delivered',
    'cancelled'
  )),

  design_approved_at TIMESTAMPTZ,
  delivery_deadline DATE,
  shipped_at TIMESTAMPTZ,
  delivered_at TIMESTAMPTZ,

  driver_or_shipping TEXT,
  approved_design_file TEXT,
  is_template_only BOOLEAN DEFAULT FALSE,
  notes TEXT,
  original_email TEXT,

  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW(),

  UNIQUE(project_id, order_number)
);

CREATE INDEX IF NOT EXISTS idx_alut_orders_project ON public.alut_orders(project_id);
CREATE INDEX IF NOT EXISTS idx_alut_orders_status ON public.alut_orders(status);
CREATE INDEX IF NOT EXISTS idx_alut_orders_deadline ON public.alut_orders(delivery_deadline);

CREATE TABLE IF NOT EXISTS public.alut_status_history (
  id SERIAL PRIMARY KEY,
  order_id INTEGER REFERENCES public.alut_orders(id) ON DELETE CASCADE,
  from_status TEXT,
  to_status TEXT NOT NULL,
  changed_by TEXT DEFAULT 'system',
  notes TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_alut_history_order ON public.alut_status_history(order_id);

CREATE OR REPLACE FUNCTION public.calculate_alut_deadline()
RETURNS TRIGGER AS $$
DECLARE
  business_days_added INTEGER := 0;
  current_date_check DATE;
BEGIN
  IF NEW.design_approved_at IS NOT NULL
     AND (OLD.design_approved_at IS NULL OR OLD.design_approved_at != NEW.design_approved_at) THEN

    current_date_check := NEW.design_approved_at::DATE;

    WHILE business_days_added < 14 LOOP
      current_date_check := current_date_check + INTERVAL '1 day';
      IF EXTRACT(DOW FROM current_date_check) NOT IN (5, 6) THEN
        business_days_added := business_days_added + 1;
      END IF;
    END LOOP;

    NEW.delivery_deadline := current_date_check;
  END IF;

  NEW.updated_at := NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_alut_deadline ON public.alut_orders;
CREATE TRIGGER trg_alut_deadline
BEFORE UPDATE ON public.alut_orders
FOR EACH ROW
EXECUTE FUNCTION public.calculate_alut_deadline();

CREATE OR REPLACE FUNCTION public.log_alut_status_change()
RETURNS TRIGGER AS $$
BEGIN
  IF OLD.status IS DISTINCT FROM NEW.status THEN
    INSERT INTO public.alut_status_history (order_id, from_status, to_status)
    VALUES (NEW.id, OLD.status, NEW.status);
  END IF;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_alut_status_log ON public.alut_orders;
CREATE TRIGGER trg_alut_status_log
AFTER UPDATE ON public.alut_orders
FOR EACH ROW
EXECUTE FUNCTION public.log_alut_status_change();

ALTER TABLE public.alut_pricing ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.alut_projects ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.alut_orders ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.alut_status_history ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "alut_pricing_all" ON public.alut_pricing;
DROP POLICY IF EXISTS "alut_projects_all" ON public.alut_projects;
DROP POLICY IF EXISTS "alut_orders_all" ON public.alut_orders;
DROP POLICY IF EXISTS "alut_history_all" ON public.alut_status_history;

CREATE POLICY "alut_pricing_all" ON public.alut_pricing FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "alut_projects_all" ON public.alut_projects FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "alut_orders_all" ON public.alut_orders FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "alut_history_all" ON public.alut_status_history FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.alut_pricing TO anon, authenticated;
GRANT ALL ON public.alut_projects TO anon, authenticated;
GRANT ALL ON public.alut_orders TO anon, authenticated;
GRANT ALL ON public.alut_status_history TO anon, authenticated;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
