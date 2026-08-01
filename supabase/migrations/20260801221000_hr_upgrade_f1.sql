-- ============================================================
-- F1 - HR upgrade: employees HR fields + seed of current staff
--
-- IMPORTANT: run this manually in the Supabase SQL editor.
-- Existing employee rows are UPDATED in place, matched by the
-- full_name that is currently in the live table (UUIDs and names
-- are NOT changed). Only employees that don't exist are inserted.
--
-- Live-name mapping (live table -> package list):
--   'איציק צמח'  = 'יצחק צמח'
--   'ליהיא אורון' = 'ליהיא אורן'
--   'ולדה'        = 'ולדה שומיקו'
-- ============================================================

-- 1) New HR columns (backwards compatible - all nullable)
ALTER TABLE public.employees
  ADD COLUMN IF NOT EXISTS id_number TEXT,
  ADD COLUMN IF NOT EXISTS address TEXT,
  ADD COLUMN IF NOT EXISTS role_title TEXT,
  ADD COLUMN IF NOT EXISTS work_start_date DATE,
  ADD COLUMN IF NOT EXISTS salary_type TEXT CHECK (salary_type IN ('monthly', 'hourly', 'monthly_global')),
  ADD COLUMN IF NOT EXISTS base_salary NUMERIC,
  ADD COLUMN IF NOT EXISTS email TEXT,
  ADD COLUMN IF NOT EXISTS pension_provider TEXT,
  ADD COLUMN IF NOT EXISTS section_14_status TEXT CHECK (section_14_status IN ('yes_full', 'yes_cap', 'no', 'pending')),
  ADD COLUMN IF NOT EXISTS employment_type TEXT CHECK (employment_type IN ('shareholder', 'employee', 'contractor')),
  ADD COLUMN IF NOT EXISTS last_salary_review_date DATE,
  ADD COLUMN IF NOT EXISTS last_contract_review_date DATE;

-- 2) Indexes
CREATE INDEX IF NOT EXISTS idx_employees_role ON public.employees(role_title);
CREATE INDEX IF NOT EXISTS idx_employees_start_date ON public.employees(work_start_date);
CREATE INDEX IF NOT EXISTS idx_employees_last_review ON public.employees(last_salary_review_date);

-- 3) Update the 9 employees that already exist in the live table.
--    Matched by their CURRENT full_name -> keeps existing UUIDs,
--    names, phones, birth dates and linked documents untouched.
UPDATE public.employees e SET
  id_number = v.id_number,
  address = NULLIF(v.address, ''),
  role_title = v.role_title,
  work_start_date = v.work_start_date::date,
  salary_type = v.salary_type,
  base_salary = v.base_salary,
  pension_provider = v.pension_provider,
  section_14_status = v.section_14_status,
  employment_type = v.employment_type,
  updated_at = NOW()
FROM (VALUES
  ('כפיר צמח',    '200334530', 'תירוש 19, שהם',              'מנכ"ל',                    '2015-01-01', 'monthly_global', 30730, 'הפניקס',           'yes_cap', 'shareholder'),
  ('נטלי צמח',    '300363314', 'שבזי 32, ראש העין',          'סמנכ"לית',                 '2012-06-01', 'monthly_global', 21685, 'הפניקס + השתלמות', 'yes_cap', 'shareholder'),
  ('איציק צמח',   '55629232',  'עופרה חזה 4, ראש העין',      'עובד ייצור',               '2015-01-01', 'monthly_global', 17620, 'קרן השתלמות',      'no',      'employee'),
  ('איריס צמח',   '57653503',  'עופרה חזה 4, ראש העין',      'מנהלת משרד',               '2016-04-01', 'monthly_global', 8800,  'קרן השתלמות',      'no',      'employee'),
  ('דניאל צמח',   '205433444', 'יונתן רטוש 1, ראש העין',     'דפָּס - מכונות דפוס',       '2017-01-01', 'monthly_global', 15948, 'הפניקס + הראל',    'yes_cap', 'employee'),
  ('ברק צמח',     '204690705', 'סביונים 6',                  'מנהל ביגוד/נהג/מכירות',    '2015-01-01', 'monthly_global', 15622, 'הפניקס + הראל',    'yes_cap', 'employee'),
  ('ליהיא אורון', '204614382', 'נחמה 26, ראשון לציון',       'מעצבת גרפית',              '2019-02-03', 'monthly_global', 8245,  'כלל פנסיה',        'yes_cap', 'employee'),
  ('אדם שולמן',   '332590041', 'שלמה בן יוסף 18, תל אביב',   'דפָּס - מכונה דיגיטלית',    '2025-01-01', 'monthly_global', 13550, 'הראל פנסיה',       'yes_cap', 'employee'),
  ('ולדה',        '341148179', '',                           'עובדת ביגוד/הדפסה',        '2025-07-16', 'hourly',         45,    'הפניקס',           'yes_cap', 'employee')
) AS v(match_name, id_number, address, role_title, work_start_date, salary_type, base_salary, pension_provider, section_14_status, employment_type)
WHERE e.full_name = v.match_name;

-- 4) Insert only the employees that don't exist yet (guarded by name AND id number)
INSERT INTO public.employees
  (full_name, id_number, address, role_title, work_start_date, salary_type, base_salary, pension_provider, section_14_status, employment_type, is_active)
SELECT v.full_name, v.id_number, NULLIF(v.address, ''), v.role_title, v.work_start_date::date,
       v.salary_type, v.base_salary, v.pension_provider, v.section_14_status, v.employment_type, true
FROM (VALUES
  ('מישל פושמיאנסקי',  '341037901', 'החשמונאים 46, בת ים',        'עובדת ייצור כללית',       '2026-06-29', 'hourly', 45, 'בהמתנה',        'pending', 'employee'),
  ('אלכסיי אובסניקוב', '034593647', 'הקיבוצים 61 דירה 1, חיפה',   'דפָּס דיגיטלי/פלוטרים',   '2026-08-02', 'hourly', 60, 'אחרי 3 חודשים', 'yes_cap', 'employee')
) AS v(full_name, id_number, address, role_title, work_start_date, salary_type, base_salary, pension_provider, section_14_status, employment_type)
WHERE NOT EXISTS (
  SELECT 1 FROM public.employees e
  WHERE e.full_name = v.full_name
     OR (e.id_number IS NOT NULL AND e.id_number = v.id_number)
);

-- 5) Unique index on id number (after data is in place)
CREATE UNIQUE INDEX IF NOT EXISTS idx_employees_id_number_unique
  ON public.employees(id_number)
  WHERE id_number IS NOT NULL AND id_number != '';

NOTIFY pgrst, 'reload schema';
