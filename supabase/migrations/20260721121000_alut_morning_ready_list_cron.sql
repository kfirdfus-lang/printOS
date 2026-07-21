-- Daily morning ready-list for Karin at 09:00 Israel (06:00 UTC)
do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname = 'alut-morning-ready-list'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'alut-morning-ready-list',
    '0 6 * * *',
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/alut-morning-ready-list',
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
