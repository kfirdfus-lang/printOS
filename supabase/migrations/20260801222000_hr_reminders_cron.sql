-- HR reminders check - 1st day of every month at 09:00 Israel (06:00 UTC)
-- Run manually in the Supabase SQL editor (requires pg_cron + pg_net, already used by other jobs).
do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname = 'hr-reminders-check'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'hr-reminders-check',
    '0 6 1 * *',
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/hr-reminders-check',
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
