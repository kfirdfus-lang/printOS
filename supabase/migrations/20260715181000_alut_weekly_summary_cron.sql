-- Weekly Alut summary - Every Sunday at 08:00 Israel (05:00 UTC)
do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname = 'alut-weekly-summary'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'alut-weekly-summary',
    '0 5 * * 0',  -- Sunday 05:00 UTC = 08:00 Israel
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/alut-weekly-summary',
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
