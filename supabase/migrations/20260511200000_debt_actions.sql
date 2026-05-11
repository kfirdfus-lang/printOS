-- supabase/migrations/20260511200000_debt_actions.sql
-- טבלת פעולות מעקב על חובות לקוחות

CREATE TABLE IF NOT EXISTS debt_actions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  
  -- זיהוי הלקוח
  bina_customer_id text NOT NULL,
  client_id uuid REFERENCES clients(id) ON DELETE SET NULL,
  client_name text,
  
  -- פעולה
  action_type text NOT NULL,  -- 'called' / 'email_sent' / 'whatsapp_sent' / 'promised_to_pay' / 
                                -- 'paid_credit' / 'paid_check' / 'paid_transfer' / 'paid_partial' / 'debt_lost'
  
  -- פרטים נוספים (תלוי בסוג הפעולה)
  notes text,                       -- הערות חופשיות
  amount numeric(12, 2),            -- סכום (לתשלום חלקי / מלא)
  
  -- להבטחה לשלם
  promised_date date,               -- מתי הבטיח לשלם
  remind_at date,                   -- מתי להזכיר חזרה אם לא שילם
  promised_amount numeric(12, 2),   -- כמה הבטיח (אופציונלי)
  
  -- לתשלום בצ'ק
  check_number text,                -- מספר צ'ק
  check_date date,                  -- תאריך צ'ק
  
  -- מטא
  created_at timestamptz DEFAULT now(),
  created_by text,                  -- מי ביצע את הפעולה (admin name)
  
  -- האם הפעולה "פתוחה" (למשל הבטחה שעוד לא הסתיימה)
  is_resolved boolean DEFAULT false,
  resolved_at timestamptz
);

-- אינדקסים
CREATE INDEX IF NOT EXISTS idx_debt_actions_customer 
  ON debt_actions(bina_customer_id);

CREATE INDEX IF NOT EXISTS idx_debt_actions_type 
  ON debt_actions(action_type);

CREATE INDEX IF NOT EXISTS idx_debt_actions_remind 
  ON debt_actions(remind_at)
  WHERE remind_at IS NOT NULL AND is_resolved = false;

CREATE INDEX IF NOT EXISTS idx_debt_actions_created_at 
  ON debt_actions(created_at DESC);

-- View נוח לדשבורד - מציג את הפעולה האחרונה לכל לקוח
CREATE OR REPLACE VIEW latest_debt_action_per_client AS
SELECT DISTINCT ON (bina_customer_id) 
  bina_customer_id,
  action_type,
  notes,
  amount,
  promised_date,
  remind_at,
  is_resolved,
  created_at,
  created_by
FROM debt_actions
ORDER BY bina_customer_id, created_at DESC;

-- View לתזכורות שצריך לקבל
CREATE OR REPLACE VIEW debt_reminders_due AS
SELECT 
  d.id,
  d.bina_customer_id,
  d.client_name,
  c.name AS current_client_name,
  d.promised_date,
  d.remind_at,
  d.promised_amount,
  d.notes,
  d.created_at,
  d.created_by,
  (CURRENT_DATE - d.remind_at) AS days_overdue
FROM debt_actions d
LEFT JOIN clients c ON c.bina_customer_id = d.bina_customer_id
WHERE d.action_type = 'promised_to_pay'
  AND d.is_resolved = false
  AND d.remind_at <= CURRENT_DATE
ORDER BY d.remind_at ASC;

COMMENT ON TABLE debt_actions IS 'מעקב פעולות גבייה - שיחות, מיילים, הבטחות, תשלומים';
COMMENT ON COLUMN debt_actions.action_type IS 
  'called=📞 / email_sent=📧 / whatsapp_sent=💬 / promised_to_pay=🤝 / paid_credit=💳 / paid_check=🏦 / paid_transfer=💸 / paid_partial=💰 / debt_lost=❌';
