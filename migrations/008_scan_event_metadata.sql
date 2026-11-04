-- Keep production scan events aligned with the shared-collection API.
-- Safe to run on databases where these columns already exist.
ALTER TABLE public.cf_scan_events
  ADD COLUMN IF NOT EXISTS event_type text NOT NULL DEFAULT 'scan';

ALTER TABLE public.cf_scan_events
  ADD COLUMN IF NOT EXISTS product_description text;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM pg_constraint
    WHERE conrelid = 'public.cf_scan_events'::regclass
      AND conname = 'cf_scan_events_event_type_check'
  ) THEN
    ALTER TABLE public.cf_scan_events
      ADD CONSTRAINT cf_scan_events_event_type_check
      CHECK (event_type IN ('scan', 'adjustment'));
  END IF;
END
$$;
