-- G3: daily cron — restore snoozed Gmail messages to INBOX.
-- Run in SQL editor (do not db push).

do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname = 'gmail-snooze-release'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'gmail-snooze-release',
    '0 5 * * *',
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/gmail-snooze-release',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'x-cron-secret', coalesce((
        select decrypted_secret::text
        from vault.decrypted_secrets
        where name = 'printos_cron_secret'
        limit 1
      ), '')
    ),
    body := '{}'::jsonb,
    timeout_milliseconds := 120000
) as request_id;
    $cmd$
  );
end
$sched$;
