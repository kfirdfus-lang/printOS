-- תעודות משלוח + חתימה דיגיטלית (מספור עצמאי מ-7566)

CREATE SEQUENCE IF NOT EXISTS delivery_note_number_seq
  START WITH 7566
  INCREMENT BY 1;

CREATE TABLE IF NOT EXISTS public.delivery_notes (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  delivery_number integer UNIQUE NOT NULL DEFAULT nextval('delivery_note_number_seq'),
  task_id uuid REFERENCES public.tasks(id) ON DELETE SET NULL,
  bina_order_id text,
  customer_name text NOT NULL,
  contact_name text,
  customer_address text,
  customer_phone text,
  customer_email text,
  items jsonb NOT NULL DEFAULT '[]'::jsonb,
  signature_token text UNIQUE,
  signature_token_expires_at timestamptz,
  signature_data text,
  signed_at timestamptz,
  signed_by_name text,
  status text NOT NULL DEFAULT 'created'
    CHECK (status IN ('created', 'sent', 'signed', 'delivered')),
  notes text,
  driver_name text,
  created_at timestamptz DEFAULT now(),
  created_by text,
  updated_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_delivery_notes_task ON public.delivery_notes(task_id);
CREATE INDEX IF NOT EXISTS idx_delivery_notes_number ON public.delivery_notes(delivery_number);
CREATE INDEX IF NOT EXISTS idx_delivery_notes_status ON public.delivery_notes(status);
CREATE INDEX IF NOT EXISTS idx_delivery_notes_token ON public.delivery_notes(signature_token)
  WHERE signature_token IS NOT NULL;

ALTER TABLE public.delivery_notes ENABLE ROW LEVEL SECURITY;

CREATE POLICY "delivery_notes_select" ON public.delivery_notes
  FOR SELECT USING (true);

CREATE POLICY "delivery_notes_insert" ON public.delivery_notes
  FOR INSERT WITH CHECK (true);

CREATE POLICY "delivery_notes_update" ON public.delivery_notes
  FOR UPDATE USING (true) WITH CHECK (true);

GRANT SELECT, INSERT, UPDATE ON public.delivery_notes TO anon, authenticated;
GRANT USAGE, SELECT ON SEQUENCE delivery_note_number_seq TO anon, authenticated;
