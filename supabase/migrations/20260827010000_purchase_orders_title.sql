-- Package L fix — optional title on purchase orders
alter table public.purchase_orders
  add column if not exists title text;

comment on column public.purchase_orders.title is
  'כותרת אופציונלית להזמנת רכש (תצוגה, PDF, שם קובץ)';
