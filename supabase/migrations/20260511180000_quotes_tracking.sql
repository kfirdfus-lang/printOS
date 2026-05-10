-- supabase/migrations/20260511180000_quotes_tracking.sql
-- מעקב אוטומטי אחר הצעות - האם נסגרו, תקועות, פג תוקף

ALTER TABLE quotes
  ADD COLUMN IF NOT EXISTS auto_status text DEFAULT 'ממתינה',
  ADD COLUMN IF NOT EXISTS closed_at timestamptz,
  ADD COLUMN IF NOT EXISTS closed_by_task_id uuid REFERENCES tasks(id) ON DELETE SET NULL,
  ADD COLUMN IF NOT EXISTS total_amount numeric(12, 2),
  ADD COLUMN IF NOT EXISTS cutoff_date timestamptz DEFAULT now(),
  ADD COLUMN IF NOT EXISTS last_status_check timestamptz;

-- הצעות "ישנות" שלא נעקוב אחריהן בדשבורד החדש
-- (575, 576 וכל מה שכבר קיים)
UPDATE quotes 
SET auto_status = 'קודמת',
    cutoff_date = now()
WHERE auto_status = 'ממתינה' OR auto_status IS NULL;

-- אינדקסים
CREATE INDEX IF NOT EXISTS idx_quotes_auto_status 
  ON quotes(auto_status);

CREATE INDEX IF NOT EXISTS idx_quotes_active 
  ON quotes(created_at DESC) 
  WHERE auto_status IN ('ממתינה', 'תקועה');

CREATE INDEX IF NOT EXISTS idx_quotes_closed_at 
  ON quotes(closed_at DESC) 
  WHERE auto_status = 'נסגרה';

-- View נוח לדשבורד
CREATE OR REPLACE VIEW quotes_dashboard AS
SELECT 
  q.id,
  q.title,
  q.bina_doc_id,
  q.bina_cust_id,
  q.bina_cust_name,
  q.contact_person,
  q.sales_agent,
  q.total_amount,
  q.auto_status,
  q.created_at,
  q.closed_at,
  q.closed_by_task_id,
  -- ימים מהיצירה
  EXTRACT(EPOCH FROM (NOW() - q.created_at)) / 86400 AS days_since_created,
  -- ימים עד פגת תוקף (30 יום מהיצירה)
  EXTRACT(EPOCH FROM ((q.created_at + INTERVAL '30 days') - NOW())) / 86400 AS days_until_expiry
FROM quotes q
WHERE q.auto_status != 'קודמת';

COMMENT ON COLUMN quotes.auto_status IS 'סטטוס אוטומטי: ממתינה / תקועה / נסגרה / פגת_תוקף / קודמת';
COMMENT ON COLUMN quotes.closed_by_task_id IS 'ה-task ב-bina שגרם לסגירת ההצעה (זוהה אוטומטית)';
COMMENT ON COLUMN quotes.cutoff_date IS 'תאריך התחלת מעקב - הצעות שנוצרו לפני זה לא מופיעות בדשבורד';
