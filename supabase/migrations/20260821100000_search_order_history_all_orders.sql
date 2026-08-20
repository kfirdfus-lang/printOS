-- Package I §7 fix: history search includes all Bina orders (archive + active)
create or replace function public.search_order_history(
  p_q text default null,
  p_from date default null,
  p_to date default null,
  p_department text default null,
  p_amount_min numeric default null,
  p_amount_max numeric default null,
  p_sort text default 'date_desc',
  p_limit int default 50,
  p_offset int default 0
)
returns jsonb
language plpgsql
stable
as $$
declare
  v_q text := nullif(trim(both from coalesce(p_q, '')), '');
  v_q_like text;
  v_limit int := greatest(1, least(coalesce(p_limit, 50), 100));
  v_offset int := greatest(0, coalesce(p_offset, 0));
  v_dept text := nullif(trim(both from coalesce(p_department, '')), '');
  v_sort text := coalesce(nullif(trim(both from coalesce(p_sort, '')), ''), 'date_desc');
  v_total int;
  v_rows jsonb;
begin
  if v_q is not null then
    v_q_like := '%' || replace(replace(replace(v_q, '\', '\\'), '%', '\%'), '_', '\_') || '%';
  end if;

  with matched as (
    select
      t.id,
      t.bina_order_id,
      t.client_name,
      t.title,
      t.bina_order_date,
      t.total_amount,
      t.total_inc_vat,
      t.discount_amount,
      t.sales_agent,
      t.contact,
      t.dept,
      t.bina_cust_id,
      t.created_at,
      t.bina_order_status,
      t.is_archive,
      coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) as sort_date,
      coalesce(t.total_inc_vat, t.total_amount, 0) as sort_amount
    from public.tasks t
    where t.bina_order_id is not null
      and (p_from is null or coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) >= p_from)
      and (p_to is null or coalesce(t.bina_order_date, (t.created_at at time zone 'Asia/Jerusalem')::date) <= p_to)
      and (p_amount_min is null or coalesce(t.total_inc_vat, t.total_amount, 0) >= p_amount_min)
      and (p_amount_max is null or coalesce(t.total_inc_vat, t.total_amount, 0) <= p_amount_max)
      and (
        v_dept is null
        or exists (
          select 1 from public.task_items ti
          where ti.task_id = t.id and ti.department = v_dept
        )
      )
      and (
        v_q is null
        or t.client_name ilike v_q_like escape '\'
        or coalesce(t.title, '') ilike v_q_like escape '\'
        or t.bina_order_id::text ilike v_q_like escape '\'
        or exists (
          select 1 from public.task_items ti
          where ti.task_id = t.id
            and coalesce(ti.description, '') ilike v_q_like escape '\'
        )
      )
  ),
  counted as (
    select count(*)::int as total from matched
  ),
  page as (
    select m.*
    from matched m
    order by
      case when v_sort = 'date_asc' then m.sort_date end asc nulls last,
      case when v_sort = 'date_desc' then m.sort_date end desc nulls last,
      case when v_sort = 'amount_asc' then m.sort_amount end asc nulls last,
      case when v_sort = 'amount_desc' then m.sort_amount end desc nulls last,
      case when v_sort = 'client_asc' then m.client_name end asc nulls last,
      m.bina_order_id desc nulls last
    limit v_limit
    offset v_offset
  ),
  enriched as (
    select
      p.id,
      p.bina_order_id,
      p.client_name,
      p.title,
      p.bina_order_date,
      p.total_amount,
      p.total_inc_vat,
      p.discount_amount,
      p.sales_agent,
      p.contact,
      p.dept,
      p.bina_cust_id,
      p.created_at,
      p.bina_order_status,
      p.is_archive,
      (
        select coalesce(jsonb_agg(
          jsonb_build_object(
            'id', ti.id,
            'description', ti.description,
            'quantity', ti.quantity,
            'price', ti.price,
            'total', ti.total,
            'department', ti.department,
            'bina_item_code', ti.bina_item_code,
            'line_number', ti.line_number
          ) order by ti.line_number nulls last
        ), '[]'::jsonb)
        from public.task_items ti
        where ti.task_id = p.id
      ) as items
    from page p
  )
  select
    (select total from counted),
    coalesce((select jsonb_agg(to_jsonb(e)) from enriched e), '[]'::jsonb)
  into v_total, v_rows;

  return jsonb_build_object(
    'total', coalesce(v_total, 0),
    'rows', coalesce(v_rows, '[]'::jsonb)
  );
end;
$$;

grant execute on function public.search_order_history(text, date, date, text, numeric, numeric, text, int, int)
  to anon, authenticated, service_role;

NOTIFY pgrst, 'reload schema';
