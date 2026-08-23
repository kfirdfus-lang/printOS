-- Package J — daily admin-tasks morning reminder at 05:00 UTC (≈08:00 Israel)
-- Requires: pg_cron, pg_net, vault secret named service_role_key
do $sched$
declare
  jid integer;
  secret_ok boolean;
begin
  select exists(
    select 1 from vault.decrypted_secrets where name = 'service_role_key' limit 1
  ) into secret_ok;

  if not secret_ok then
    raise exception 'vault secret service_role_key is missing — aborting cron schedule';
  end if;

  for jid in select jobid from cron.job where jobname = 'admin-tasks-reminder'
  loop
    perform cron.unschedule(jid);
  end loop;

  perform cron.schedule(
    'admin-tasks-reminder',
    '0 5 * * *',
    $cmd$
select net.http_post(
    url := 'https://pvwcpukfhyrmdpxgfwrk.supabase.co/functions/v1/admin-tasks-reminder',
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
