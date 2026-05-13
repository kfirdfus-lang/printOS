-- supabase/migrations/20260510120000_create_debt_snapshots.sql
-- טבלה לשמירת snapshots יומיים של דוח חייבים מבינה

CREATE TABLE IF NOT EXISTS debt_snapshots (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  
  -- מתי נשלף
  snapshot_date date NOT NULL DEFAULT CURRENT_DATE,
  fetched_at timestamptz NOT NULL DEFAULT now(),
  
  -- פרטי החשבונית מבינה
  bina_customer_id text NOT NULL,
  customer_name text,                    -- מ-clients table (אם קיים)
  doc_num integer NOT NULL,              -- מספר חשבונית
  doc_date date,                         -- תאריך חשבונית
  doc_payment_date date,                 -- תאריך תשלום צפוי (לפי תנאי תשלום)
  doc_total numeric(12, 2),              -- סכום החשבונית
  doc_balance numeric(12, 2) NOT NULL,   -- יתרה לתשלום
  
  -- חישובים
  is_overdue boolean NOT NULL DEFAULT false,  -- האם עבר תאריך תשלום
  days_overdue integer DEFAULT 0,             -- כמה ימים באיחור
  
  -- חיבור ללקוח אצלנו (אם קיים)
  client_id uuid REFERENCES clients(id) ON DELETE SET NULL,
  client_exists_in_db boolean NOT NULL DEFAULT false
);

-- אינדקסים לחיפושים מהירים
CREATE INDEX IF NOT EXISTS idx_debt_snapshots_date 
  ON debt_snapshots(snapshot_date DESC);

CREATE INDEX IF NOT EXISTS idx_debt_snapshots_customer 
  ON debt_snapshots(bina_customer_id, snapshot_date DESC);

CREATE INDEX IF NOT EXISTS idx_debt_snapshots_overdue 
  ON debt_snapshots(snapshot_date, is_overdue) 
  WHERE is_overdue = true;

CREATE INDEX IF NOT EXISTS idx_debt_snapshots_open 
  ON debt_snapshots(snapshot_date) 
  WHERE doc_balance > 0;

-- אילוץ: לא יכול להיות כפילות של אותה חשבונית באותו יום
CREATE UNIQUE INDEX IF NOT EXISTS idx_debt_snapshots_unique
  ON debt_snapshots(snapshot_date, bina_customer_id, doc_num);

-- View נוח לדשבורד מנהלים - תמונת מצב נוכחית
CREATE OR REPLACE VIEW current_debt_summary AS
SELECT 
  bina_customer_id,
  customer_name,
  client_id,
  client_exists_in_db,
  COUNT(*) AS open_invoices_count,
  SUM(doc_balance) AS total_debt,
  SUM(CASE WHEN is_overdue THEN doc_balance ELSE 0 END) AS overdue_debt,
  MAX(days_overdue) AS max_days_overdue,
  MIN(doc_payment_date) AS earliest_payment_date
FROM debt_snapshots
WHERE snapshot_date = (SELECT MAX(snapshot_date) FROM debt_snapshots)
  AND doc_balance > 0
GROUP BY bina_customer_id, customer_name, client_id, client_exists_in_db
ORDER BY total_debt DESC;

-- View לסיכום כללי
CREATE OR REPLACE VIEW current_debt_totals AS
SELECT
  snapshot_date,
  COUNT(DISTINCT bina_customer_id) AS customers_with_debt,
  COUNT(*) AS open_invoices,
  SUM(doc_balance) AS total_open_debt,
  SUM(CASE WHEN is_overdue THEN doc_balance ELSE 0 END) AS total_overdue_debt,
  COUNT(CASE WHEN is_overdue THEN 1 END) AS overdue_invoices_count,
  COUNT(DISTINCT CASE WHEN is_overdue THEN bina_customer_id END) AS overdue_customers_count
FROM debt_snapshots
WHERE snapshot_date = (SELECT MAX(snapshot_date) FROM debt_snapshots)
  AND doc_balance > 0
GROUP BY snapshot_date;

COMMENT ON TABLE debt_snapshots IS 'Snapshot יומי של חוב פתוח מבינה (docType -900). cron יומי ב-07:00.';
COMMENT ON COLUMN debt_snapshots.is_overdue IS 'doc_payment_date < היום AND doc_balance > 0';
COMMENT ON COLUMN debt_snapshots.client_exists_in_db IS 'האם הלקוח קיים בטבלת clients (לפי bina_customer_id)';
