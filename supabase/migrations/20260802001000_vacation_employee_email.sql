-- Add optional employee_email to vacation_requests (for decision notification emails).
-- Run manually in the Supabase SQL editor.
ALTER TABLE public.vacation_requests
  ADD COLUMN IF NOT EXISTS employee_email TEXT;

NOTIFY pgrst, 'reload schema';
