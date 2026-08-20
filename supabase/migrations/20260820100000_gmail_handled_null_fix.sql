-- G4: ensure handled is never null (pending filter)
update public.gmail_classifications
set handled = false
where handled is null;

NOTIFY pgrst, 'reload schema';
