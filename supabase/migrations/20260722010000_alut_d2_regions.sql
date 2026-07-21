-- Alut D2: delivery_city + city_regions mapping
ALTER TABLE public.alut_orders
  ADD COLUMN IF NOT EXISTS delivery_city TEXT;

CREATE INDEX IF NOT EXISTS idx_alut_orders_delivery_city
  ON public.alut_orders(delivery_city)
  WHERE delivery_city IS NOT NULL;

CREATE TABLE IF NOT EXISTS public.city_regions (
  id SERIAL PRIMARY KEY,
  city_name TEXT NOT NULL UNIQUE,
  region TEXT NOT NULL CHECK (region IN ('צפון', 'חיפה', 'שרון', 'מרכז', 'ירושלים', 'שפלה', 'דרום')),
  aliases TEXT[],
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_city_regions_name ON public.city_regions(city_name);
CREATE INDEX IF NOT EXISTS idx_city_regions_region ON public.city_regions(region);

ALTER TABLE public.city_regions ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "city_regions_read_all" ON public.city_regions;
CREATE POLICY "city_regions_read_all" ON public.city_regions FOR SELECT USING (true);

DROP POLICY IF EXISTS "city_regions_write" ON public.city_regions;
CREATE POLICY "city_regions_write" ON public.city_regions FOR ALL USING (true) WITH CHECK (true);

GRANT SELECT ON public.city_regions TO anon;
GRANT ALL ON public.city_regions TO authenticated;
GRANT USAGE, SELECT ON SEQUENCE public.city_regions_id_seq TO anon, authenticated;

INSERT INTO public.city_regions (city_name, region, aliases) VALUES
-- צפון
('נהריה', 'צפון', NULL),
('כרמיאל', 'צפון', NULL),
('צפת', 'צפון', NULL),
('טבריה', 'צפון', NULL),
('נצרת', 'צפון', NULL),
('בית שאן', 'צפון', NULL),
('מגדל העמק', 'צפון', NULL),
('יקנעם', 'צפון', ARRAY['יוקנעם', 'יקנעם עילית', 'יוקנעם עילית']),
('נוף הגליל', 'צפון', ARRAY['נצרת עילית']),
('קריית שמונה', 'צפון', ARRAY['קרית שמונה', 'ק. שמונה']),
('מעלות', 'צפון', ARRAY['מעלות תרשיחא']),
('כפר יאסיף', 'צפון', NULL),
('עכו', 'צפון', NULL),
('שלומי', 'צפון', NULL),
-- חיפה
('חיפה', 'חיפה', NULL),
('קריית מוצקין', 'חיפה', ARRAY['קרית מוצקין', 'ק. מוצקין']),
('קריית ביאליק', 'חיפה', ARRAY['קרית ביאליק', 'ק. ביאליק']),
('קריית ים', 'חיפה', ARRAY['קרית ים', 'ק. ים']),
('קריית אתא', 'חיפה', ARRAY['קרית אתא', 'ק. אתא']),
('טירת כרמל', 'חיפה', NULL),
('זכרון יעקב', 'חיפה', ARRAY['זיכרון יעקב', 'זכרון']),
('בנימינה', 'חיפה', ARRAY['בנימינה גבעת עדה']),
('קיסריה', 'חיפה', NULL),
('חדרה', 'חיפה', NULL),
('נשר', 'חיפה', NULL),
('פרדס חנה', 'חיפה', ARRAY['פרדס חנה כרכור']),
-- שרון
('נתניה', 'שרון', NULL),
('כפר סבא', 'שרון', ARRAY['כפ"ס', 'כפ״ס']),
('הוד השרון', 'שרון', NULL),
('רעננה', 'שרון', NULL),
('הרצליה', 'שרון', NULL),
('רמת השרון', 'שרון', NULL),
('אבן יהודה', 'שרון', NULL),
('קדימה', 'שרון', ARRAY['קדימה צורן', 'צורן']),
('גני תקווה', 'שרון', NULL),
('חוף השרון', 'שרון', NULL),
('רמת הכובש', 'שרון', NULL),
('כפר יונה', 'שרון', NULL),
('אלישמע', 'שרון', NULL),
-- מרכז
('תל אביב', 'מרכז', ARRAY['ת"א', 'ת״א', 'תל אביב יפו', 'תל-אביב']),
('רמת גן', 'מרכז', ARRAY['ר"ג', 'ר״ג']),
('גבעתיים', 'מרכז', NULL),
('בני ברק', 'מרכז', ARRAY['ב"ב', 'ב״ב']),
('פתח תקווה', 'מרכז', ARRAY['פ"ת', 'פ״ת', 'פתח תקוה']),
('ראש העין', 'מרכז', NULL),
('יהוד', 'מרכז', ARRAY['יהוד מונוסון']),
('אור יהודה', 'מרכז', NULL),
('קריית אונו', 'מרכז', ARRAY['קרית אונו', 'ק. אונו']),
('סביון', 'מרכז', NULL),
-- ירושלים
('ירושלים', 'ירושלים', ARRAY['ים', 'י-ם']),
('מבשרת ציון', 'ירושלים', ARRAY['מבשרת']),
('מעלה אדומים', 'ירושלים', NULL),
('בית שמש', 'ירושלים', NULL),
('גבעת זאב', 'ירושלים', NULL),
('אפרת', 'ירושלים', NULL),
('אלעד', 'ירושלים', NULL),
('ביתר עילית', 'ירושלים', NULL),
-- שפלה
('ראשון לציון', 'שפלה', ARRAY['ראש"ל', 'ראשל"צ', 'ראשון']),
('חולון', 'שפלה', NULL),
('בת ים', 'שפלה', NULL),
('רחובות', 'שפלה', NULL),
('נס ציונה', 'שפלה', NULL),
('מודיעין', 'שפלה', ARRAY['מודיעין מכבים רעות']),
('לוד', 'שפלה', NULL),
('רמלה', 'שפלה', NULL),
('קריית עקרון', 'שפלה', ARRAY['קרית עקרון', 'ק. עקרון']),
('אזור', 'שפלה', NULL),
('גדרה', 'שפלה', NULL),
('יבנה', 'שפלה', NULL),
-- דרום
('אשדוד', 'דרום', NULL),
('אשקלון', 'דרום', NULL),
('קריית מלאכי', 'דרום', ARRAY['קרית מלאכי', 'ק. מלאכי']),
('קריית גת', 'דרום', ARRAY['קרית גת', 'ק. גת']),
('שדרות', 'דרום', NULL),
('נתיבות', 'דרום', NULL),
('אופקים', 'דרום', NULL),
('באר שבע', 'דרום', ARRAY['ב"ש', 'ב״ש']),
('דימונה', 'דרום', NULL),
('ערד', 'דרום', NULL),
('אילת', 'דרום', NULL),
('מצפה רמון', 'דרום', NULL),
('מיתר', 'דרום', NULL),
('להבים', 'דרום', NULL),
('עומר', 'דרום', NULL)
ON CONFLICT (city_name) DO NOTHING;

NOTIFY pgrst, 'reload schema';
