-- supabase/migrations/20260512000000_clients_emails_separation.sql
-- הפרדת מיילים לגבייה / הזמנות מוכנות

ALTER TABLE clients
  ADD COLUMN IF NOT EXISTS collection_email_primary text,
  ADD COLUMN IF NOT EXISTS collection_email_secondary text,
  ADD COLUMN IF NOT EXISTS orders_email_primary text,
  ADD COLUMN IF NOT EXISTS orders_email_secondary text;

COMMENT ON COLUMN clients.collection_email_primary IS 'מייל ראשי לגבייה';
COMMENT ON COLUMN clients.collection_email_secondary IS 'מייל משני לגבייה';
COMMENT ON COLUMN clients.orders_email_primary IS 'מייל ראשי לעדכוני הזמנות מוכנות';
COMMENT ON COLUMN clients.orders_email_secondary IS 'מייל משני לעדכוני הזמנות מוכנות';
