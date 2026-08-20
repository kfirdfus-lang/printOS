-- G4: Gmail UI preferences on users (signature auto-append, push toggles later)
alter table public.users
  add column if not exists gmail_auto_signature boolean not null default true;

NOTIFY pgrst, 'reload schema';
