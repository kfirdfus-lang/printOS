ALTER TABLE clients
  ADD COLUMN IF NOT EXISTS contact_name text;
COMMENT ON COLUMN clients.contact_name IS 'שם איש קשר לגבייה';
