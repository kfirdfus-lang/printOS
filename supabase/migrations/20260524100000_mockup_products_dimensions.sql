-- מימדים אמיתיים למוצרי הדמיה (הכנה לקנבס אינטראקטיבי)
ALTER TABLE public.mockup_products
  ADD COLUMN IF NOT EXISTS real_width_cm numeric DEFAULT 40,
  ADD COLUMN IF NOT EXISTS real_height_cm numeric DEFAULT 50;

UPDATE public.mockup_products SET real_width_cm = 50, real_height_cm = 70
WHERE category = 'clothing' AND name LIKE '%T%';

UPDATE public.mockup_products SET real_width_cm = 52, real_height_cm = 75
WHERE category = 'clothing' AND (name LIKE '%שרוול ארוך%' OR name LIKE '%פולו%');

UPDATE public.mockup_products SET real_width_cm = 55, real_height_cm = 75
WHERE category = 'clothing' AND (name LIKE '%קפוצ%' OR name LIKE '%וסט%');

UPDATE public.mockup_products SET real_width_cm = 25, real_height_cm = 30
WHERE category = 'clothing' AND name LIKE '%תינוק%';

UPDATE public.mockup_products SET real_width_cm = 22, real_height_cm = 10
WHERE category = 'clothing' AND name LIKE '%כובע%';

UPDATE public.mockup_products SET real_width_cm = 38, real_height_cm = 42
WHERE category = 'accessories' AND name LIKE '%תיק%';

UPDATE public.mockup_products SET real_width_cm = 10, real_height_cm = 12
WHERE category = 'drinkware';

UPDATE public.mockup_products SET real_width_cm = 21, real_height_cm = 29.7
WHERE category = 'marketing' AND name LIKE '%A4%';

UPDATE public.mockup_products SET real_width_cm = 14.8, real_height_cm = 21
WHERE category = 'marketing' AND name LIKE '%A5%';

UPDATE public.mockup_products SET real_width_cm = 10.5, real_height_cm = 14.8
WHERE category = 'marketing' AND name LIKE '%A6%';

UPDATE public.mockup_products SET real_width_cm = 9, real_height_cm = 5
WHERE category = 'marketing' AND name LIKE '%ביקור%';

UPDATE public.mockup_products SET real_width_cm = 85, real_height_cm = 200
WHERE category = 'displays' AND name LIKE '%רולאפ%';

UPDATE public.mockup_products SET real_width_cm = 80, real_height_cm = 120
WHERE category = 'displays' AND name LIKE '%מולטי%';

UPDATE public.mockup_products SET real_width_cm = 70, real_height_cm = 100
WHERE category = 'displays' AND name LIKE '%חמור%';

UPDATE public.mockup_products SET real_width_cm = 160, real_height_cm = 60
WHERE category = 'displays' AND name LIKE '%מגנטי%';

NOTIFY pgrst, 'reload schema';
