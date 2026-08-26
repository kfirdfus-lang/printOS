-- Package K2 — get_customer_stats (server-side aggregates)
-- Match by bina_cust_id when provided, else by client_name (case-insensitive trim).

create or replace function public.get_customer_stats(
  p_client_name text default null,
  p_bina_cust_id bigint default null
)
returns jsonb
language plpgsql
stable
security invoker
as $$
declare
  v_name text := nullif(trim(both from coalesce(p_client_name, '')), '');
  v_result jsonb;
begin
  if p_bina_cust_id is null and v_name is null then
    return jsonb_build_object(
      'total_orders', 0,
      'total_amount', 0,
      'avg_order', 0,
      'first_order_date', null,
      'last_order_date', null,
      'days_since_last', null,
      'avg_days_between', null,
      'monthly', '[]'::jsonb,
      'departments', '[]'::jsonb,
      'top_items', '[]'::jsonb,
      'recent_orders', '[]'::jsonb
    );
  end if;

  with base as (
    select
      t.id,
      t.bina_order_id,
      t.client_name,
      t.title,
      t.bina_order_date,
      t.total_amount,
      t.total_inc_vat,
      t.sales_agent,
      t.bina_cust_id,
      t.created_at,
      coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) as order_date,
      coalesce(t.total_amount, 0)::numeric as amt
    from public.tasks t
    where t.bina_order_id is not null
      and (
        (p_bina_cust_id is not null and t.bina_cust_id = p_bina_cust_id)
        or (
          p_bina_cust_id is null
          and v_name is not null
          and lower(trim(both from coalesce(t.client_name, ''))) = lower(v_name)
        )
      )
  ),
  ordered as (
    select *, row_number() over (order by order_date asc, bina_order_id asc) as rn
    from base
  ),
  gaps as (
    select
      (o.order_date - lag(o.order_date) over (order by o.order_date asc, o.bina_order_id asc))::numeric as gap_days
    from ordered o
  ),
  summary as (
    select
      count(*)::int as total_orders,
      coalesce(sum(amt), 0)::numeric as total_amount,
      case when count(*) > 0 then round(coalesce(sum(amt), 0) / count(*), 2) else 0 end as avg_order,
      min(order_date) as first_order_date,
      max(order_date) as last_order_date,
      case
        when max(order_date) is null then null
        else (current_date - max(order_date))::int
      end as days_since_last,
      (
        select round(avg(g.gap_days)::numeric, 1)
        from gaps g
        where g.gap_days is not null and g.gap_days >= 0
      ) as avg_days_between
    from base
  ),
  month_spine as (
    select to_char(d, 'YYYY-MM') as month
    from generate_series(
      date_trunc('month', current_date) - interval '11 months',
      date_trunc('month', current_date),
      interval '1 month'
    ) as d
  ),
  monthly_agg as (
    select to_char(date_trunc('month', order_date), 'YYYY-MM') as month,
           count(*)::int as orders,
           coalesce(sum(amt), 0)::numeric as amount
    from base
    where order_date >= (date_trunc('month', current_date) - interval '11 months')::date
    group by 1
  ),
  monthly as (
    select jsonb_agg(
      jsonb_build_object(
        'month', s.month,
        'orders', coalesce(a.orders, 0),
        'amount', coalesce(a.amount, 0)
      )
      order by s.month
    ) as arr
    from month_spine s
    left join monthly_agg a on a.month = s.month
  ),
  dept_agg as (
    select
      ti.department,
      count(distinct ti.task_id)::int as orders,
      coalesce(sum(ti.total), 0)::numeric as amount
    from public.task_items ti
    join base b on b.id = ti.task_id
    where ti.department is not null and trim(ti.department) <> ''
    group by ti.department
  ),
  departments as (
    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'department', d.department,
          'orders', d.orders,
          'amount', d.amount
        )
        order by d.amount desc
      ),
      '[]'::jsonb
    ) as arr
    from dept_agg d
  ),
  item_agg as (
    select
      trim(ti.description) as description,
      count(*)::int as times,
      coalesce(sum(ti.quantity), 0)::numeric as total_qty,
      coalesce(sum(ti.total), 0)::numeric as total_sum,
      coalesce(sum(ti.price), 0)::numeric as price_sum
    from public.task_items ti
    join base b on b.id = ti.task_id
    where ti.description is not null
      and trim(ti.description) <> ''
      and trim(ti.description) <> '—'
      and trim(ti.description) <> '-'
      and trim(ti.description) <> '–'
    group by trim(ti.description)
    having count(*) >= 2
      and not (
        coalesce(sum(ti.quantity), 0) = 0
        and coalesce(sum(ti.price), 0) = 0
        and coalesce(sum(ti.total), 0) = 0
      )
  ),
  top_items as (
    select coalesce(
      (
        select jsonb_agg(
          jsonb_build_object(
            'description', i.description,
            'times', i.times,
            'total_qty', i.total_qty
          )
          order by i.times desc, i.total_qty desc
        )
        from (
          select * from item_agg
          order by times desc, total_qty desc
          limit 15
        ) i
      ),
      '[]'::jsonb
    ) as arr
  ),
  item_money as (
    select
      coalesce(sum(ti.total), 0)::numeric as all_items_amount,
      coalesce(
        sum(ti.total) filter (
          where ti.department is not null and trim(ti.department) <> ''
        ),
        0
      )::numeric as dept_assigned_amount
    from public.task_items ti
    join base b on b.id = ti.task_id
  ),
  recent as (
    select
      b.id,
      b.bina_order_id,
      b.title,
      b.order_date,
      b.amt as total_amount,
      b.total_inc_vat,
      b.sales_agent,
      coalesce(
        (
          select jsonb_agg(
            jsonb_build_object(
              'description', ti.description,
              'quantity', ti.quantity,
              'price', ti.price,
              'total', ti.total,
              'department', ti.department,
              'bina_item_code', ti.bina_item_code,
              'line_number', ti.line_number
            )
            order by ti.line_number nulls last
          )
          from public.task_items ti
          where ti.task_id = b.id
        ),
        '[]'::jsonb
      ) as items
    from base b
    order by b.order_date desc nulls last, b.bina_order_id desc nulls last
    limit 20
  ),
  recent_orders as (
    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', r.id,
          'bina_order_id', r.bina_order_id,
          'title', r.title,
          'order_date', r.order_date,
          'total_amount', r.total_amount,
          'total_inc_vat', r.total_inc_vat,
          'sales_agent', r.sales_agent,
          'items', r.items
        )
      ),
      '[]'::jsonb
    ) as arr
    from recent r
  )
  select jsonb_build_object(
    'total_orders', s.total_orders,
    'total_amount', s.total_amount,
    'avg_order', s.avg_order,
    'first_order_date', s.first_order_date,
    'last_order_date', s.last_order_date,
    'days_since_last', s.days_since_last,
    'avg_days_between', s.avg_days_between,
    'monthly', coalesce(m.arr, '[]'::jsonb),
    'departments', coalesce(d.arr, '[]'::jsonb),
    'top_items', coalesce(t.arr, '[]'::jsonb),
    'recent_orders', coalesce(ro.arr, '[]'::jsonb),
    'dept_assigned_amount', coalesce(im.dept_assigned_amount, 0),
    'all_items_amount', coalesce(im.all_items_amount, 0)
  )
  into v_result
  from summary s
  cross join monthly m
  cross join departments d
  cross join top_items t
  cross join recent_orders ro
  cross join item_money im;

  return coalesce(v_result, '{}'::jsonb);
end;
$$;

grant execute on function public.get_customer_stats(text, bigint) to anon, authenticated, service_role;

comment on function public.get_customer_stats(text, bigint) is
  'Package K2: customer order stats for client card. Amounts from tasks.total_amount (Bina orders, not invoices).';
