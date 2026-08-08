-- Package E v1.1 — supplier invoice parsing (parse-only, no Bina API).
-- Run manually in the Supabase SQL editor.

-- ------------------------------------------------------------
-- 1. Supplier mapping (tax ID / name → Bina supplier code)
-- ------------------------------------------------------------
create table if not exists public.bina_suppliers (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),

  bina_supplier_code text not null,
  supplier_name text not null,
  supplier_tax_id text,

  address text,
  city text,
  phone text,
  fax text,
  email text,

  default_expense_type text default 'עסקי',
  default_payment_terms text,

  is_active boolean not null default true,
  notes text,

  constraint bina_suppliers_code_unique unique (bina_supplier_code)
);

create index if not exists idx_bina_suppliers_tax_id
  on public.bina_suppliers (supplier_tax_id);
create index if not exists idx_bina_suppliers_name
  on public.bina_suppliers (supplier_name);

alter table public.bina_suppliers enable row level security;

drop policy if exists "bina_suppliers_all" on public.bina_suppliers;
create policy "bina_suppliers_all"
  on public.bina_suppliers for all using (true) with check (true);

grant all on public.bina_suppliers to anon, authenticated;

comment on table public.bina_suppliers is
  'חבילה E — מיפוי ספקים: ח.פ מהחשבונית ← קוד ספק בבינה';


-- ------------------------------------------------------------
-- 2. Parsed supplier invoices
-- ------------------------------------------------------------
create table if not exists public.supplier_invoices (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  created_by_name text,

  file_path text,
  file_name text,
  file_type text,
  file_size_kb integer,

  parse_status text not null default 'pending'
    check (parse_status in ('pending','success','failed')),
  parse_error text,
  model_used text,
  parse_duration_ms integer,
  raw_response jsonb,

  supplier_name text,
  supplier_tax_id text,
  supplier_address text,
  supplier_city text,
  supplier_phone text,
  matched_supplier_id uuid references public.bina_suppliers(id) on delete set null,
  bina_supplier_code text,
  match_method text,

  allocation_number text,
  invoice_number text,
  invoice_date date,
  vat_date date,
  payment_terms text,
  due_date date,
  expense_type text default 'עסקי',
  currency text default 'ILS',

  amount_before_vat numeric(12,2),
  discount_percent numeric(5,2),
  discount_amount numeric(12,2),
  amount_after_discount numeric(12,2),
  vat_rate numeric(5,2),
  vat_amount numeric(12,2),
  total_amount numeric(12,2),

  line_items jsonb default '[]'::jsonb,

  low_confidence_fields jsonb default '[]'::jsonb,
  ai_notes text,

  parse_quality text
    check (parse_quality in ('perfect','minor_errors','major_errors')),
  quality_notes text,
  reviewed_at timestamptz,

  entered_in_bina boolean not null default false,
  entered_in_bina_at timestamptz,

  notes text
);

create index if not exists idx_supplier_invoices_created_at
  on public.supplier_invoices (created_at desc);
create index if not exists idx_supplier_invoices_supplier
  on public.supplier_invoices (supplier_name);
create index if not exists idx_supplier_invoices_tax_id
  on public.supplier_invoices (supplier_tax_id);
create index if not exists idx_supplier_invoices_date
  on public.supplier_invoices (invoice_date desc);
create index if not exists idx_supplier_invoices_status
  on public.supplier_invoices (parse_status);
create index if not exists idx_supplier_invoices_quality
  on public.supplier_invoices (parse_quality);

alter table public.supplier_invoices enable row level security;

drop policy if exists "supplier_invoices_all" on public.supplier_invoices;
create policy "supplier_invoices_all"
  on public.supplier_invoices for all using (true) with check (true);

grant all on public.supplier_invoices to anon, authenticated;

comment on table public.supplier_invoices is
  'חבילה E — חשבוניות ספקים מפורסרות ב-AI. שלב 1: פרסור בלבד, ללא אינטגרציה עם בינה.';


-- ------------------------------------------------------------
-- 3. Seed first supplier (from sample invoice)
-- ------------------------------------------------------------
insert into public.bina_suppliers
  (bina_supplier_code, supplier_name, supplier_tax_id, address, city, phone, fax)
values
  ('210', 'פונגר 2000 בע"מ', '510895634', 'לוינסקי 140', 'תל אביב', '03-6880319', '03-6880938')
on conflict (bina_supplier_code) do nothing;

notify pgrst, 'reload schema';
