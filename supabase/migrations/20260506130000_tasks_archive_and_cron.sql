-- Task archive + pg_cron → Edge Functions (archive-completed-tasks / cleanup-old-archive).
--
-- 1) Vault (Dashboard → Project Settings → Vault): create secret named printos_cron_secret
--    with the SAME value as Edge Function secrets CRON_SECRET on both archive + cleanup functions.
-- 2) Daily job: cron 0 0 * * * = ~03:00 Israel (IDT, UTC+3). Adjust in cron.job if you need exactly 03:00 year-round.
-- 3) Weekly cleanup: cron 0 0 * * 0 = Sundays 00:00 UTC (~03:00 IDT).

alter table public.tasks
  add column if not exists archived_at timestamptz,
  add column if not exists completed_at timestamptz,
  add column if not exists updated_at timestamptz default now();

comment on column public.tasks.archived_at is 'When non-null, task is in admin archive (hidden from main board).';

update public.tasks
set updated_at = coalesce(updated_at, created_at, now())
where updated_at is null;

create or replace function public.tasks_set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at := now();
  return new;
end;
$$;

drop trigger if exists tasks_set_updated_at on public.tasks;
create trigger tasks_set_updated_at
  before update on public.tasks
  for each row
  execute function public.tasks_set_updated_at();

-- Cron + HTTP client (usually available on hosted Supabase; ignore if unsupported locally).
create extension if not exists pg_cron with schema pg_catalog;
create extension if not exists pg_net with schema extensions;

do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname in ('printos_archive_completed_daily', 'printos_cleanup_old_archive_weekly')
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'printos_archive_completed_daily',
    '0 0 * * *',
    $cmd$
select net.http_post(
  url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/archive-completed-tasks',
  headers := jsonb_build_object(
    'Content-Type', 'application/json',
    'x-cron-secret', coalesce((select decrypted_secret::text from vault.decrypted_secrets where name = 'printos_cron_secret' limit 1), '')
  ),
  body := '{}'::jsonb
);
    $cmd$
  );

  perform cron.schedule(
    'printos_cleanup_old_archive_weekly',
    '0 0 * * 0',
    $cmd$
select net.http_post(
  url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/cleanup-old-archive',
  headers := jsonb_build_object(
    'Content-Type', 'application/json',
    'x-cron-secret', coalesce((select decrypted_secret::text from vault.decrypted_secrets where name = 'printos_cron_secret' limit 1), '')
  ),
  body := '{}'::jsonb
);
    $cmd$
  );
end
$sched$;
