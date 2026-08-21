-- Package I: server-side quote history search (archive + active with bina_doc_id)
create or replace function public.search_quote_history(
  p_q text default null,
  p_from date default null,
  p_to date default null,
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
  v_sort text := coalesce(nullif(trim(both from coalesce(p_sort, '')), ''), 'date_desc');
  v_total int;
  v_rows jsonb;
begin
  if v_q is not null then
    v_q_like := '%' || replace(replace(replace(v_q, '\', '\\'), '%', '\%'), '_', '\_') || '%';
  end if;

  with matched as (
    select
      q.id,
      q.bina_doc_id,
      q.bina_cust_id,
      q.bina_cust_name,
      q.title,
      q.bina_doc_date,
      q.total_amount,
      q.total,
      q.subtotal,
      q.vat_amount,
      q.sales_agent,
      q.contact_person,
      q.status,
      q.quote_status,
      q.sent_at,
      q.created_at,
      q.is_archive,
      coalesce(q.bina_doc_date, (coalesce(q.sent_at, q.created_at) at time zone 'Asia/Jerusalem')::date) as sort_date,
      coalesce(q.total_amount, q.total, q.subtotal, 0) as sort_amount
    from public.quotes q
    where q.bina_doc_id is not null
      and q.deleted_at is null
      and (p_from is null or coalesce(q.bina_doc_date, (coalesce(q.sent_at, q.created_at) at time zone 'Asia/Jerusalem')::date) >= p_from)
      and (p_to is null or coalesce(q.bina_doc_date, (coalesce(q.sent_at, q.created_at) at time zone 'Asia/Jerusalem')::date) <= p_to)
      and (p_amount_min is null or coalesce(q.total_amount, q.total, q.subtotal, 0) >= p_amount_min)
      and (p_amount_max is null or coalesce(q.total_amount, q.total, q.subtotal, 0) <= p_amount_max)
      and (
        v_q is null
        or q.bina_cust_name ilike v_q_like escape '\'
        or coalesce(q.title, '') ilike v_q_like escape '\'
        or q.bina_doc_id::text ilike v_q_like escape '\'
        or exists (
          select 1 from public.quote_items qi
          where qi.quote_id = q.id
            and coalesce(qi.description, '') ilike v_q_like escape '\'
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
      case when v_sort = 'client_asc' then m.bina_cust_name end asc nulls last,
      m.bina_doc_id desc nulls last
    limit v_limit
    offset v_offset
  ),
  enriched as (
    select
      p.id,
      p.bina_doc_id,
      p.bina_cust_id,
      p.bina_cust_name,
      p.title,
      p.bina_doc_date,
      p.total_amount,
      p.total,
      p.subtotal,
      p.vat_amount,
      p.sales_agent,
      p.contact_person,
      p.status,
      p.quote_status,
      p.sent_at,
      p.created_at,
      p.is_archive,
      (
        select coalesce(jsonb_agg(
          jsonb_build_object(
            'id', qi.id,
            'description', qi.description,
            'quantity', qi.quantity,
            'unit_price', qi.unit_price,
            'discount_pct', qi.discount_pct,
            'total', qi.total,
            'item_name', qi.item_name,
            'line_number', qi.line_number
          ) order by qi.line_number nulls last
        ), '[]'::jsonb)
        from public.quote_items qi
        where qi.quote_id = p.id
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

grant execute on function public.search_quote_history(text, date, date, numeric, numeric, text, int, int)
  to anon, authenticated, service_role;

NOTIFY pgrst, 'reload schema';
