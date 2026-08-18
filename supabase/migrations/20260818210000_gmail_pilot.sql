-- Gmail pilot (read-only). Tokens never leave the service role.
-- Run in the Supabase SQL editor if db push is blocked.

CREATE TABLE IF NOT EXISTS public.gmail_oauth_states (
  state TEXT PRIMARY KEY,
  user_id TEXT NOT NULL REFERENCES public.users(id) ON DELETE CASCADE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  expires_at TIMESTAMPTZ NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_gmail_oauth_states_expires
  ON public.gmail_oauth_states (expires_at);

CREATE TABLE IF NOT EXISTS public.gmail_connections (
  user_id TEXT PRIMARY KEY REFERENCES public.users(id) ON DELETE CASCADE,
  google_email TEXT,
  access_token TEXT NOT NULL,
  refresh_token TEXT,
  token_expiry TIMESTAMPTZ,
  scope TEXT,
  connected_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE public.gmail_oauth_states ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.gmail_connections ENABLE ROW LEVEL SECURITY;

REVOKE ALL ON public.gmail_oauth_states FROM anon, authenticated;
REVOKE ALL ON public.gmail_connections FROM anon, authenticated;

COMMENT ON TABLE public.gmail_connections IS
  'Gmail OAuth tokens for the read-only pilot. Service-role access only. No send/delete/modify.';

NOTIFY pgrst, 'reload schema';
