-- Persist private participant drafts so the same user can resume on another device.
-- Apply once to the dedicated ColetaFlow Neon database using the database owner.
CREATE TABLE IF NOT EXISTS cf_participant_draft_events (
  event_id uuid PRIMARY KEY,
  collection_id uuid NOT NULL REFERENCES cf_collections(id) ON DELETE CASCADE,
  participant_id uuid NOT NULL REFERENCES cf_participants(id) ON DELETE CASCADE,
  barcode text NOT NULL,
  quantity numeric(14,3) NOT NULL CHECK (quantity <> 0),
  event_type text NOT NULL DEFAULT 'scan' CHECK (event_type IN ('scan','adjustment')),
  device_id text,
  scanned_at timestamptz NOT NULL,
  product_description text,
  saved_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS cf_participant_draft_events_participant_idx
  ON cf_participant_draft_events(collection_id,participant_id,saved_at);
