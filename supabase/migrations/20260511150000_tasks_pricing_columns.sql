-- Pricing / sales fields synced from Bina orders + dashboard support
ALTER TABLE public.tasks
  ADD COLUMN IF NOT EXISTS total_amount numeric(14, 2),
  ADD COLUMN IF NOT EXISTS total_inc_vat numeric(14, 2),
  ADD COLUMN IF NOT EXISTS discount_amount numeric(14, 2),
  ADD COLUMN IF NOT EXISTS sales_agent text,
  ADD COLUMN IF NOT EXISTS bina_order_date date;

COMMENT ON COLUMN public.tasks.total_amount IS 'סכום לפני מע״מ / אחרי הנחה — מבינה (orderTotalAfterDiscount / orderTotal)';
COMMENT ON COLUMN public.tasks.total_inc_vat IS 'סכום כולל מע״מ — מבינה (orderTotalIncVat)';
COMMENT ON COLUMN public.tasks.discount_amount IS 'סכום הנחה — מבינה (orderDiscount)';
COMMENT ON COLUMN public.tasks.sales_agent IS 'סוכן מכירות מבינה (orderSalesMan)';
COMMENT ON COLUMN public.tasks.bina_order_date IS 'תאריך הזמנה בבינה (orderDate, DD/MM/YYYY → date)';
