-- Package H Stage B — daily digest for open quotes older than 14 days
-- 06:00 UTC ≈ 09:00 Israel (standard time). Run manually in SQL editor.
do $sched$
declare
  jid integer;
begin
  for jid in select jobid from cron.job where jobname = 'quotes-follow-up'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'quotes-follow-up',
    '0 6 * * *',
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/quotes-follow-up',
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
