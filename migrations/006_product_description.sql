-- Add optional product description to shared scan events.
-- Apply once using the Neon database owner/migration role.
-- Runtime application credentials must not perform schema changes.
ALTER TABLE cf_scan_events
  ADD COLUMN IF NOT EXISTS product_description text;
