-- ============================================================
-- חבילה L — הזמנות רכש
-- ============================================================

create sequence if not exists public.purchase_order_seq start with 8001;

create table if not exists public.purchase_orders (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  created_by text,

  po_number integer not null default nextval('public.purchase_order_seq'),

  supplier_id uuid references public.bina_suppliers(id) on delete set null,
  supplier_name text not null,
  supplier_contact text,

  po_date date not null default current_date,
  expected_date date,

  status text not null default 'draft'
    check (status in ('draft','sent','received','cancelled')),

  payment_terms text,
  notes text,

  subtotal numeric(12,2) default 0,
  vat_rate numeric(5,2) default 18,
  vat_amount numeric(12,2) default 0,
  total numeric(12,2) default 0,

  received_at timestamptz,
  received_by text,

  updated_at timestamptz not null default now(),

  constraint purchase_orders_number_unique unique (po_number)
);

create index if not exists idx_po_status on public.purchase_orders (status);
create index if not exists idx_po_date on public.purchase_orders (po_date desc);
create index if not exists idx_po_supplier on public.purchase_orders (supplier_id);

alter table public.purchase_orders enable row level security;
drop policy if exists "purchase_orders_all" on public.purchase_orders;
create policy "purchase_orders_all"
  on public.purchase_orders for all using (true) with check (true);


create table if not exists public.purchase_order_items (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),

  po_id uuid not null references public.purchase_orders(id) on delete cascade,
  line_number integer,

  description text not null,
  quantity numeric(12,2) default 1,
  unit text,
  unit_price numeric(12,4),
  price_entry_mode text default 'unit'
    check (price_entry_mode in ('unit','total')),
  total numeric(12,2)
);

create index if not exists idx_po_items_po
  on public.purchase_order_items (po_id, line_number);

alter table public.purchase_order_items enable row level security;
drop policy if exists "purchase_order_items_all" on public.purchase_order_items;
create policy "purchase_order_items_all"
  on public.purchase_order_items for all using (true) with check (true);

grant all on public.purchase_orders to anon, authenticated, service_role;
grant all on public.purchase_order_items to anon, authenticated, service_role;
grant usage, select on sequence public.purchase_order_seq to anon, authenticated, service_role;
