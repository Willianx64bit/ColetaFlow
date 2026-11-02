-- Preserve every scan and explicit quantity correction as an independent event.
-- Apply only to the dedicated ColetaFlow database.
BEGIN;
ALTER TABLE cf_scan_events ADD COLUMN IF NOT EXISTS event_type text NOT NULL DEFAULT 'scan';
ALTER TABLE cf_scan_events DROP CONSTRAINT IF EXISTS cf_scan_events_quantity_check;
ALTER TABLE cf_scan_events ADD CONSTRAINT cf_scan_events_quantity_nonzero_check CHECK (quantity <> 0);
ALTER TABLE cf_scan_events ADD CONSTRAINT cf_scan_events_event_type_check CHECK (event_type IN ('scan','adjustment'));
CREATE OR REPLACE VIEW cf_collection_totals AS
SELECT collection_id,barcode,SUM(quantity) AS total_quantity,
       COUNT(*) AS scan_count,MAX(received_at) AS last_received_at
FROM cf_scan_events
GROUP BY collection_id,barcode;
COMMIT;
