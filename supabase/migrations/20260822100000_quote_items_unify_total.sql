-- Unify quote_items amount column to `total` (canonical).
-- `line_total` was added later for parity with order_items but quote_items
-- already had NOT NULL `total` used by import + older code.

-- 1) Move any values that only exist on line_total
update public.quote_items
set total = line_total
where line_total is not null
  and (total is null or total = 0)
  and line_total <> 0;

-- Prefer line_total when both differ (should be rare / none)
update public.quote_items
set total = line_total
where line_total is not null
  and total is distinct from line_total;

-- 2) Drop duplicate column
alter table public.quote_items
  drop column if exists line_total;

-- 3) Ensure columns the UI needs exist (idempotent)
alter table public.quote_items
  add column if not exists bina_item_code text,
  add column if not exists department text;

NOTIFY pgrst, 'reload schema';
