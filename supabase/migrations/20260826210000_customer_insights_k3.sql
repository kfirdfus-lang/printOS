-- Package K3 — insights RPCs + dormant dismiss flags

alter table public.clients
  add column if not exists dormant_dismissed boolean not null default false,
  add column if not exists dormant_dismissed_at timestamptz;

-- ---------------------------------------------------------------------------
-- Dormant customers: >=3 orders in last year, last order 90+ days ago
-- Sorted by amount DESC (not recency)
-- ---------------------------------------------------------------------------
create or replace function public.get_dormant_customers()
returns jsonb
language plpgsql
stable
security invoker
as $$
declare
  v_out jsonb;
begin
  with orders as (
    select
      t.client_name,
      t.bina_cust_id,
      coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) as order_date,
      coalesce(t.total_amount, 0)::numeric as amt
    from public.tasks t
    where t.bina_order_id is not null
      and coalesce(t.client_name, '') <> ''
  ),
  last_year as (
    select *
    from orders
    where order_date >= (current_date - interval '365 days')::date
  ),
  by_client as (
    select
      o.client_name,
      max(o.bina_cust_id) as bina_cust_id,
      count(*) filter (where o.order_date >= (current_date - interval '365 days')::date)::int as orders_last_year,
      coalesce(sum(o.amt) filter (where o.order_date >= (current_date - interval '365 days')::date), 0)::numeric as amount_last_year,
      max(o.order_date) as last_order_date,
      (current_date - max(o.order_date))::int as days_since_last
    from orders o
    group by o.client_name
  ),
  dormant as (
    select
      b.*,
      (
        select ti.department
        from public.task_items ti
        join public.tasks t on t.id = ti.task_id
        where t.bina_order_id is not null
          and lower(trim(t.client_name)) = lower(trim(b.client_name))
          and ti.department is not null and trim(ti.department) <> ''
        group by ti.department
        order by count(*) desc
        limit 1
      ) as top_department
    from by_client b
    where b.orders_last_year >= 3
      and b.days_since_last >= 90
      and not exists (
        select 1 from public.clients c
        where c.dormant_dismissed = true
          and (
            (b.bina_cust_id is not null and c.bina_customer_id::text = b.bina_cust_id::text)
            or lower(trim(c.name)) = lower(trim(b.client_name))
          )
      )
  )
  select coalesce(
    jsonb_agg(
      jsonb_build_object(
        'client_name', d.client_name,
        'bina_cust_id', d.bina_cust_id,
        'orders_last_year', d.orders_last_year,
        'amount_last_year', d.amount_last_year,
        'last_order_date', d.last_order_date,
        'days_since_last', d.days_since_last,
        'top_department', d.top_department
      )
      order by d.amount_last_year desc nulls last
    ),
    '[]'::jsonb
  )
  into v_out
  from dormant d;

  return coalesce(v_out, '[]'::jsonb);
end;
$$;

grant execute on function public.get_dormant_customers() to anon, authenticated, service_role;

-- ---------------------------------------------------------------------------
-- Seasonality — all months present in DB
-- ---------------------------------------------------------------------------
create or replace function public.get_seasonality_stats()
returns jsonb
language plpgsql
stable
security invoker
as $$
declare
  v_out jsonb;
begin
  with base as (
    select
      coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) as order_date,
      coalesce(t.total_amount, 0)::numeric as amt
    from public.tasks t
    where t.bina_order_id is not null
  ),
  bounds as (
    select min(order_date) as first_d, max(order_date) as last_d,
           count(distinct date_trunc('month', order_date))::int as months_with_data
    from base
  ),
  by_month as (
    select extract(month from order_date)::int as month_num,
           count(*)::int as orders,
           coalesce(sum(amt), 0)::numeric as amount
    from base
    group by 1
  )
  select jsonb_build_object(
    'months_with_data', b.months_with_data,
    'first_date', b.first_d,
    'last_date', b.last_d,
    'partial', coalesce(b.months_with_data, 0) < 12,
    'months', coalesce(
      (
        select jsonb_agg(
          jsonb_build_object(
            'month', m.month_num,
            'orders', coalesce(x.orders, 0),
            'amount', coalesce(x.amount, 0)
          )
          order by m.month_num
        )
        from generate_series(1, 12) as m(month_num)
        left join by_month x on x.month_num = m.month_num
      ),
      '[]'::jsonb
    )
  )
  into v_out
  from bounds b;

  return coalesce(v_out, '{}'::jsonb);
end;
$$;

grant execute on function public.get_seasonality_stats() to anon, authenticated, service_role;

-- ---------------------------------------------------------------------------
-- Price trends — normalized descriptions, 3+ occurrences
-- ---------------------------------------------------------------------------
create or replace function public.normalize_item_desc(p text)
returns text
language sql
immutable
as $$
  select nullif(
    trim(both from regexp_replace(
      regexp_replace(
        regexp_replace(
          lower(coalesce(p, '')),
          '[''\"״׳`]',
          '',
          'g'
        ),
        -- unify size separators only (/ \ *) — do NOT treat letter x as separator
        '[/\\\\*×]',
        '/',
        'g'
      ),
      '\s+',
      ' ',
      'g'
    )),
    ''
  );
$$;

create or replace function public.get_price_trends()
returns jsonb
language plpgsql
stable
security invoker
as $$
declare
  v_out jsonb;
begin
  with lines as (
    select
      public.normalize_item_desc(ti.description) as norm,
      trim(ti.description) as sample_desc,
      coalesce(ti.price, 0)::numeric as price,
      coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) as order_date
    from public.task_items ti
    join public.tasks t on t.id = ti.task_id
    where t.bina_order_id is not null
      and public.normalize_item_desc(ti.description) is not null
      and public.normalize_item_desc(ti.description) not in ('—', '-', '–')
      and coalesce(ti.price, 0) > 0
      and coalesce(ti.quantity, 0) > 0
  ),
  counted as (
    select norm, count(*)::int as times,
           min(order_date) as first_d,
           max(order_date) as last_d,
           (array_agg(sample_desc order by order_date desc))[1] as label
    from lines
    group by norm
    having count(*) >= 3
  ),
  series as (
    select
      c.norm,
      c.label,
      c.times,
      c.first_d,
      c.last_d,
      (
        select jsonb_agg(
          jsonb_build_object('date', s.order_date, 'price', s.price)
          order by s.order_date
        )
        from (
          select distinct on (date_trunc('month', l.order_date))
            l.order_date, l.price
          from lines l
          where l.norm = c.norm
          order by date_trunc('month', l.order_date), l.order_date desc
        ) s
      ) as points,
      (
        select l.price from lines l
        where l.norm = c.norm
        order by l.order_date desc
        limit 1
      ) as last_price,
      (
        select max(l.order_date) from lines l
        where l.norm = c.norm
          and l.price is distinct from (
            select l2.price from lines l2
            where l2.norm = c.norm
            order by l2.order_date desc
            limit 1
          )
      ) as last_price_change_date
    from counted c
  )
  select coalesce(
    jsonb_agg(
      jsonb_build_object(
        'key', s.norm,
        'description', s.label,
        'times', s.times,
        'points', coalesce(s.points, '[]'::jsonb),
        'last_price', s.last_price,
        'months_unchanged', case
          when s.last_price_change_date is null then
            greatest(0, (extract(year from age(current_date, s.first_d)) * 12
              + extract(month from age(current_date, s.first_d)))::int)
          else
            greatest(0, (extract(year from age(current_date, s.last_price_change_date)) * 12
              + extract(month from age(current_date, s.last_price_change_date)))::int)
        end,
        'stale', case
          when s.last_price_change_date is null then
            (extract(year from age(current_date, s.first_d)) * 12
              + extract(month from age(current_date, s.first_d))) >= 6
          else
            (extract(year from age(current_date, s.last_price_change_date)) * 12
              + extract(month from age(current_date, s.last_price_change_date))) >= 6
        end
      )
      order by s.times desc
    ),
    '[]'::jsonb
  )
  into v_out
  from series s;

  return coalesce(v_out, '[]'::jsonb);
end;
$$;

grant execute on function public.get_price_trends() to anon, authenticated, service_role;
grant execute on function public.normalize_item_desc(text) to anon, authenticated, service_role;

-- ---------------------------------------------------------------------------
-- Lead time by department (completed orders with both dates)
-- One row per (task, department) to avoid multi-counting line items
-- ---------------------------------------------------------------------------
create or replace function public.get_lead_time_by_dept()
returns jsonb
language plpgsql
stable
security invoker
as $$
declare
  v_out jsonb;
begin
  with task_depts as (
    select distinct
      t.id as task_id,
      coalesce(
        nullif(trim(ti.department), ''),
        nullif(trim(t.dept), ''),
        'לא ידוע'
      ) as department,
      (
        (t.completed_at at time zone 'Asia/Jerusalem')::date
        - coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date)
      )::numeric as lead_days
    from public.tasks t
    left join public.task_items ti on ti.task_id = t.id
    where t.bina_order_id is not null
      and t.completed_at is not null
      and coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) is not null
  ),
  filtered as (
    select * from task_depts
    where lead_days is not null and lead_days >= 0 and lead_days < 120
  ),
  agg as (
    select department,
           round(avg(lead_days)::numeric, 1) as avg_days,
           count(*)::int as sample_size
    from filtered
    group by department
    having count(*) >= 3
  )
  select coalesce(
    jsonb_agg(
      jsonb_build_object(
        'department', a.department,
        'avg_days', a.avg_days,
        'sample_size', a.sample_size,
        'slow', a.avg_days >= 5
      )
      order by a.avg_days desc
    ),
    '[]'::jsonb
  )
  into v_out
  from agg a;

  return coalesce(v_out, '[]'::jsonb);
end;
$$;

grant execute on function public.get_lead_time_by_dept() to anon, authenticated, service_role;
