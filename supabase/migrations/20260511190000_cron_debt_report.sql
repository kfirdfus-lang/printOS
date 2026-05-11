-- Cron job יומי לרענון דוח חייבים
-- דורש: pg_cron, pg_net. ב-Vault: סוד בשם service_role_key (ערך = service role JWT) או התאמת השם.

create extension if not exists pg_cron with schema pg_catalog;
create extension if not exists pg_net with schema extensions;

do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname = 'bina-fetch-debt-report-daily'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'bina-fetch-debt-report-daily',
    '0 3 * * *',  -- כל יום ב-03:00 (שעון השרת — בדרך כלל UTC; להזיז אם צריך 03:00 ישראל)
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/bina-fetch-debt-report',
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
    timeout_milliseconds := 300000
) as request_id;
    $cmd$
  );
end
$sched$;
