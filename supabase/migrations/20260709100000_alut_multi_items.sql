-- ============================================================
-- אלוט 1.5: תמיכה במספר פריטים להזמנה
-- ============================================================

CREATE TABLE IF NOT EXISTS public.alut_orders_backup AS
SELECT * FROM public.alut_orders;

CREATE TABLE IF NOT EXISTS public.alut_order_items (
  id SERIAL PRIMARY KEY,
  order_id INTEGER NOT NULL REFERENCES public.alut_orders(id) ON DELETE CASCADE,

  calendar_type TEXT NOT NULL CHECK (calendar_type IN ('hard', 'duplex', 'wall')),
  quantity INTEGER NOT NULL,
  needs_design BOOLEAN DEFAULT TRUE,
  is_template_only BOOLEAN DEFAULT FALSE,

  unit_price NUMERIC(10,2) NOT NULL,
  design_price NUMERIC(10,2) DEFAULT 0,
  cartons_count INTEGER DEFAULT 0,
  subtotal NUMERIC(10,2) DEFAULT 0,

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

  approved_design_file TEXT,

  notes TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_alut_items_order ON public.alut_order_items(order_id);
CREATE INDEX IF NOT EXISTS idx_alut_items_status ON public.alut_order_items(status);
CREATE INDEX IF NOT EXISTS idx_alut_items_deadline ON public.alut_order_items(delivery_deadline);

INSERT INTO public.alut_order_items (
  order_id, calendar_type, quantity, needs_design, is_template_only,
  unit_price, design_price, cartons_count, subtotal,
  davach_cost, davach_sent_at, davach_received_at,
  status, design_approved_at, delivery_deadline, shipped_at, delivered_at,
  approved_design_file, notes
)
SELECT
  id, calendar_type, quantity, needs_design, is_template_only,
  unit_price, design_price, cartons_count,
  (unit_price * quantity + design_price) AS subtotal,
  davach_cost, davach_sent_at, davach_received_at,
  status, design_approved_at, delivery_deadline, shipped_at, delivered_at,
  approved_design_file, notes
FROM public.alut_orders
WHERE NOT EXISTS (
  SELECT 1 FROM public.alut_order_items
  WHERE alut_order_items.order_id = alut_orders.id
)
AND calendar_type IS NOT NULL;

ALTER TABLE public.alut_orders
  DROP COLUMN IF EXISTS calendar_type CASCADE,
  DROP COLUMN IF EXISTS quantity CASCADE,
  DROP COLUMN IF EXISTS needs_design CASCADE,
  DROP COLUMN IF EXISTS unit_price CASCADE,
  DROP COLUMN IF EXISTS design_price CASCADE,
  DROP COLUMN IF EXISTS cartons_count CASCADE,
  DROP COLUMN IF EXISTS davach_cost CASCADE,
  DROP COLUMN IF EXISTS davach_sent_at CASCADE,
  DROP COLUMN IF EXISTS davach_received_at CASCADE,
  DROP COLUMN IF EXISTS status CASCADE,
  DROP COLUMN IF EXISTS design_approved_at CASCADE,
  DROP COLUMN IF EXISTS delivery_deadline CASCADE,
  DROP COLUMN IF EXISTS shipped_at CASCADE,
  DROP COLUMN IF EXISTS delivered_at CASCADE,
  DROP COLUMN IF EXISTS approved_design_file CASCADE,
  DROP COLUMN IF EXISTS is_template_only CASCADE;

ALTER TABLE public.alut_orders
  ADD COLUMN IF NOT EXISTS shipping_price NUMERIC(10,2) DEFAULT 0,
  ADD COLUMN IF NOT EXISTS total_cartons INTEGER DEFAULT 0,
  ADD COLUMN IF NOT EXISTS driver_or_shipping TEXT;

DROP TRIGGER IF EXISTS trg_alut_deadline ON public.alut_orders;

CREATE OR REPLACE FUNCTION public.calculate_alut_item_deadline()
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

DROP TRIGGER IF EXISTS trg_alut_item_deadline ON public.alut_order_items;
CREATE TRIGGER trg_alut_item_deadline
BEFORE UPDATE ON public.alut_order_items
FOR EACH ROW
EXECUTE FUNCTION public.calculate_alut_item_deadline();

ALTER TABLE public.alut_status_history
  ADD COLUMN IF NOT EXISTS item_id INTEGER REFERENCES public.alut_order_items(id) ON DELETE CASCADE;

DROP TRIGGER IF EXISTS trg_alut_status_log ON public.alut_orders;

CREATE OR REPLACE FUNCTION public.log_alut_item_status_change()
RETURNS TRIGGER AS $$
BEGIN
  IF OLD.status IS DISTINCT FROM NEW.status THEN
    INSERT INTO public.alut_status_history (order_id, item_id, from_status, to_status)
    VALUES (NEW.order_id, NEW.id, OLD.status, NEW.status);
  END IF;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_alut_item_status_log ON public.alut_order_items;
CREATE TRIGGER trg_alut_item_status_log
AFTER UPDATE ON public.alut_order_items
FOR EACH ROW
EXECUTE FUNCTION public.log_alut_item_status_change();

CREATE OR REPLACE FUNCTION public.recalculate_alut_order_totals()
RETURNS TRIGGER AS $$
DECLARE
  v_order_id INTEGER;
  v_total_cartons INTEGER;
  v_items_total NUMERIC(10,2);
  v_shipping_price NUMERIC(10,2);
  v_pricing_shipping NUMERIC(10,2);
  v_is_pickup BOOLEAN;
BEGIN
  IF TG_OP = 'DELETE' THEN
    v_order_id := OLD.order_id;
  ELSE
    v_order_id := NEW.order_id;
  END IF;

  SELECT value INTO v_pricing_shipping
  FROM public.alut_pricing WHERE key = 'price_shipping_carton';

  SELECT is_pickup INTO v_is_pickup
  FROM public.alut_orders WHERE id = v_order_id;

  SELECT
    COALESCE(SUM(cartons_count), 0),
    COALESCE(SUM(subtotal), 0)
  INTO v_total_cartons, v_items_total
  FROM public.alut_order_items
  WHERE order_id = v_order_id;

  v_shipping_price := CASE
    WHEN v_is_pickup THEN 0
    ELSE v_total_cartons * COALESCE(v_pricing_shipping, 40)
  END;

  UPDATE public.alut_orders SET
    total_cartons = v_total_cartons,
    shipping_price = v_shipping_price,
    total_price = v_items_total + v_shipping_price,
    updated_at = NOW()
  WHERE id = v_order_id;

  RETURN COALESCE(NEW, OLD);
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_alut_recalc_totals ON public.alut_order_items;
CREATE TRIGGER trg_alut_recalc_totals
AFTER INSERT OR UPDATE OR DELETE ON public.alut_order_items
FOR EACH ROW
EXECUTE FUNCTION public.recalculate_alut_order_totals();

ALTER TABLE public.alut_order_items ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "alut_items_all" ON public.alut_order_items;
CREATE POLICY "alut_items_all" ON public.alut_order_items FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.alut_order_items TO anon, authenticated;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO anon, authenticated;

-- עדכון סיכומים להזמנות קיימות
UPDATE public.alut_orders o SET
  total_cartons = COALESCE((SELECT SUM(cartons_count) FROM public.alut_order_items i WHERE i.order_id = o.id), 0),
  total_price = COALESCE((SELECT SUM(subtotal) FROM public.alut_order_items i WHERE i.order_id = o.id), 0)
    + CASE WHEN o.is_pickup THEN 0 ELSE COALESCE((SELECT SUM(cartons_count) FROM public.alut_order_items i WHERE i.order_id = o.id), 0) * COALESCE((SELECT value FROM public.alut_pricing WHERE key = 'price_shipping_carton'), 40) END,
  shipping_price = CASE WHEN o.is_pickup THEN 0 ELSE COALESCE((SELECT SUM(cartons_count) FROM public.alut_order_items i WHERE i.order_id = o.id), 0) * COALESCE((SELECT value FROM public.alut_pricing WHERE key = 'price_shipping_carton'), 40) END;

NOTIFY pgrst, 'reload schema';
