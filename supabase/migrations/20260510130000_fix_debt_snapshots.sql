-- supabase/migrations/20260510130000_fix_debt_snapshots.sql
-- מתקן את הסכמה - מוחק את הגרסה הישנה ויוצר מחדש לפי הסכמה הנכונה

-- מחיקת ה-views (חייב לפני מחיקת הטבלה)
DROP VIEW IF EXISTS current_debt_summary CASCADE;
DROP VIEW IF EXISTS current_debt_totals CASCADE;

-- מחיקת הטבלה הישנה
DROP TABLE IF EXISTS debt_snapshots CASCADE;

-- יצירה מחדש לפי הסכמה הנכונה
CREATE TABLE debt_snapshots (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  
  snapshot_date date NOT NULL DEFAULT CURRENT_DATE,
  fetched_at timestamptz NOT NULL DEFAULT now(),
  
  bina_customer_id text NOT NULL,
  customer_name text,
  doc_num integer NOT NULL,
  doc_date date,
  doc_payment_date date,
  doc_total numeric(12, 2),
  doc_balance numeric(12, 2) NOT NULL,
  
  is_overdue boolean NOT NULL DEFAULT false,
  days_overdue integer DEFAULT 0,
  
  client_id uuid REFERENCES clients(id) ON DELETE SET NULL,
  client_exists_in_db boolean NOT NULL DEFAULT false
);

CREATE INDEX idx_debt_snapshots_date 
  ON debt_snapshots(snapshot_date DESC);

CREATE INDEX idx_debt_snapshots_customer 
  ON debt_snapshots(bina_customer_id, snapshot_date DESC);

CREATE INDEX idx_debt_snapshots_overdue 
  ON debt_snapshots(snapshot_date, is_overdue) 
  WHERE is_overdue = true;

CREATE INDEX idx_debt_snapshots_open 
  ON debt_snapshots(snapshot_date) 
  WHERE doc_balance > 0;

CREATE UNIQUE INDEX idx_debt_snapshots_unique
  ON debt_snapshots(snapshot_date, bina_customer_id, doc_num);

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
