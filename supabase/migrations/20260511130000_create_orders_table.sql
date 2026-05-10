CREATE TABLE IF NOT EXISTS orders (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  bina_doc_id integer,
  title text,
  bina_cust_id integer,
  bina_cust_name text,
  bina_cust_address text,
  bina_cust_city text,
  bina_cust_phone text,
  bina_cust_email text,
  contact_person text,
  sales_agent text,
  remark text,
  status text DEFAULT 'נשלחה',
  task_id uuid REFERENCES tasks(id) ON DELETE SET NULL,
  created_at timestamptz DEFAULT now()
);

CREATE TABLE IF NOT EXISTS order_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  order_id uuid REFERENCES orders(id) ON DELETE CASCADE,
  line_number integer,
  description text,
  quantity numeric,
  unit_price numeric,
  discount_pct numeric DEFAULT 0,
  created_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_orders_bina_doc_id ON orders(bina_doc_id);
CREATE INDEX IF NOT EXISTS idx_order_items_order_id ON order_items(order_id);
