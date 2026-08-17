-- ============================================================
-- שלב ב׳ v3 — מחלקות, תנאי תשלום, שכפול
-- הרץ ידנית ב-SQL Editor. לא supabase db push.
-- ============================================================

-- 1. תיקון טבלת המחלקות + קוד בינה
alter table public.department_settings
  add column if not exists bina_item_code text;

delete from public.department_settings where name = 'משלוחים';

insert into public.department_settings (name, color, has_rip, sort_order, bina_item_code) values
  ('פורמט רחב',                  '#62C7C2', true,  1, '8'),
  ('דיגיטלי צבעוני',             '#3EA9A4', true,  2, '3'),
  ('דיגיטלי שחור לבן',           '#6B7F92', true,  3, '4'),
  ('אופסט',                      '#0E3651', true,  4, '5'),
  ('ביגוד ומוצרי פרסום',         '#EC008C', true,  5, '2'),
  ('מתקני תצוגה ומוצרים נלווים', '#00AEEF', false, 6, '7'),
  ('עבודות חוץ',                 '#D97706', false, 7, '6')
on conflict (name) do update
  set bina_item_code = excluded.bina_item_code,
      color          = excluded.color,
      has_rip        = excluded.has_rip,
      sort_order     = excluded.sort_order;

-- 2. תנאי תשלום ברמת הלקוח
alter table public.clients
  add column if not exists default_payment_terms text,
  add column if not exists payment_terms_updated_at timestamptz;

do $$
begin
  if exists (
    select 1 from information_schema.columns
    where table_schema = 'public' and table_name = 'clients' and column_name = 'payment_terms'
  ) then
    update public.clients
      set default_payment_terms = payment_terms
    where default_payment_terms is null and payment_terms is not null;
  end if;
end $$;

-- 3. מקור המחיר ברמת הפריט + קוד מחלקה לשכפול
alter table public.order_items
  add column if not exists price_entry_mode text default 'unit',
  add column if not exists bina_item_code text,
  add column if not exists line_total numeric;

alter table public.quote_items
  add column if not exists price_entry_mode text default 'unit',
  add column if not exists bina_item_code text,
  add column if not exists line_total numeric;

do $$
begin
  if not exists (
    select 1 from pg_constraint where conname = 'order_items_price_entry_mode_check'
  ) then
    alter table public.order_items
      add constraint order_items_price_entry_mode_check
      check (price_entry_mode in ('unit','total'));
  end if;
  if not exists (
    select 1 from pg_constraint where conname = 'quote_items_price_entry_mode_check'
  ) then
    alter table public.quote_items
      add constraint quote_items_price_entry_mode_check
      check (price_entry_mode in ('unit','total'));
  end if;
end $$;

-- 4. מעקב שכפול
alter table public.tasks
  add column if not exists duplicated_from_order_id text;

notify pgrst, 'reload schema';
