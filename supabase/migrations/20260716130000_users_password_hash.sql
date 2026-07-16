-- הוספת שדה password_hash לטבלת users (אימות סיסמה ב-Login)
ALTER TABLE public.users
  ADD COLUMN IF NOT EXISTS password_hash TEXT,
  ADD COLUMN IF NOT EXISTS password_updated_at TIMESTAMPTZ;

CREATE INDEX IF NOT EXISTS idx_users_password_hash ON public.users(password_hash)
  WHERE password_hash IS NOT NULL;

NOTIFY pgrst, 'reload schema';
