-- Alut D1: sent_to_natalie + ready_for_pickup statuses
ALTER TABLE public.alut_order_items
  DROP CONSTRAINT IF EXISTS alut_order_items_status_check;

ALTER TABLE public.alut_order_items
  ADD CONSTRAINT alut_order_items_status_check
  CHECK (status IN (
    'new',
    'design_in_progress',
    'design_sent',
    'design_approved',
    'sent_to_natalie',
    'in_print',
    'at_davach',
    'at_itzik',
    'ready_to_ship',
    'ready_for_pickup',
    'shipped',
    'delivered',
    'cancelled'
  ));

ALTER TABLE public.alut_order_items
  ADD COLUMN IF NOT EXISTS sent_to_natalie_at TIMESTAMPTZ;

NOTIFY pgrst, 'reload schema';
