-- Daily Alut deadline alert (08:00 Israel ≈ 05:00 UTC)
-- Requires: pg_cron, pg_net, vault secret named service_role_key

create extension if not exists pg_cron with schema pg_catalog;
create extension if not exists pg_net with schema extensions;

do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname = 'alut-deadline-alert-daily'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'alut-deadline-alert-daily',
    '0 5 * * *',
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/alut-deadline-alert',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'Authorization', 'Bearer ' || (
        select decrypted_secret::text
        from vault.decrypted_secrets
        where name = 'service_role_key'
        limit 1
      )
    ),
    body := '{}'::jsonb,
    timeout_milliseconds := 120000
) as request_id;
    $cmd$
  );
end
$sched$;
