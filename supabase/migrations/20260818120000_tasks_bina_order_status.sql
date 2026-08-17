-- סטטוס הזמנה מבינה (orderStatus / orderState) — לניתוח חשבונית ומחזור חיים
alter table public.tasks
  add column if not exists bina_order_status text,
  add column if not exists bina_order_state text;

comment on column public.tasks.bina_order_status is 'סטטוס הזמנה מבינה (orderStatus)';
comment on column public.tasks.bina_order_state is 'מצב הזמנה מבינה (orderState)';

notify pgrst, 'reload schema';
