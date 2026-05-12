-- טבלה לעקוב אחר תשלומים שזיהינו ושלחנו עליהם מייל למורן

CREATE TABLE IF NOT EXISTS detected_payments (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  bina_customer_id text NOT NULL,
  customer_name text NOT NULL,
  doc_num text NOT NULL,
  amount numeric(12,2) NOT NULL,
  doc_date timestamptz,
  doc_payment_date timestamptz,
  detected_at timestamptz NOT NULL DEFAULT now(),
  email_sent_at timestamptz,
  email_status text DEFAULT 'pending',
  email_error text,
  CONSTRAINT unique_payment_detection UNIQUE (bina_customer_id, doc_num)
);

CREATE INDEX IF NOT EXISTS idx_detected_payments_detected_at ON detected_payments(detected_at DESC);
CREATE INDEX IF NOT EXISTS idx_detected_payments_customer ON detected_payments(bina_customer_id);

COMMENT ON TABLE detected_payments IS 'תשלומים שזוהו על ידי השוואת snapshots ונשלחו על כך מייל';
COMMENT ON COLUMN detected_payments.doc_num IS 'מספר החשבונית שסולקה';
COMMENT ON COLUMN detected_payments.amount IS 'הסכום ששולם (מה-snapshot הקודם)';
COMMENT ON COLUMN detected_payments.email_status IS 'pending | sent | failed';
