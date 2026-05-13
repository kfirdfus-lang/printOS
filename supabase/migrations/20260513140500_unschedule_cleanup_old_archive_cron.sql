-- Stop automatic deletion of archived tasks (cleanup-old-archive Edge Function).
-- Keeps printos_archive_completed_daily (archive-completed-tasks) unchanged.

DO $body$
DECLARE
  r RECORD;
BEGIN
  FOR r IN
    SELECT jobid, jobname
    FROM cron.job
    WHERE command ILIKE '%cleanup-old-archive%'
       OR jobname = 'printos_cleanup_old_archive_weekly'
  LOOP
    PERFORM cron.unschedule(r.jobid);
    RAISE NOTICE 'Unscheduled cron job % (jobid %)', r.jobname, r.jobid;
  END LOOP;
END
$body$;
