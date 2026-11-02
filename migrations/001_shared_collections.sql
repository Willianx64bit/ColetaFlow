-- ColetaFlow shared collections (independent of MobFlow)
-- Apply only to the dedicated ColetaFlow Neon database.
CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS cf_collections (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  access_code text NOT NULL UNIQUE,
  title text NOT NULL DEFAULT 'Coleta compartilhada',
  status text NOT NULL DEFAULT 'open' CHECK (status IN ('open','paused','completed','archived')),
  template_file_name text,
  template_mime_type text,
  template_bytes bytea,
  template_sha256 text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  completed_at timestamptz
);

CREATE TABLE IF NOT EXISTS cf_participants (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  collection_id uuid NOT NULL REFERENCES cf_collections(id) ON DELETE CASCADE,
  display_name text NOT NULL,
  token_hash text NOT NULL UNIQUE,
  role text NOT NULL DEFAULT 'collector' CHECK (role IN ('manager','collector')),
  status text NOT NULL DEFAULT 'active' CHECK (status IN ('active','revoked','finished')),
  joined_at timestamptz NOT NULL DEFAULT now(),
  last_seen_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (collection_id, display_name)
);

CREATE TABLE IF NOT EXISTS cf_scan_events (
  event_id uuid PRIMARY KEY,
  collection_id uuid NOT NULL REFERENCES cf_collections(id) ON DELETE CASCADE,
  participant_id uuid NOT NULL REFERENCES cf_participants(id) ON DELETE RESTRICT,
  barcode text NOT NULL,
  quantity numeric(14,3) NOT NULL CHECK (quantity > 0),
  device_id text,
  scanned_at timestamptz NOT NULL,
  received_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS cf_scan_events_collection_barcode_idx
  ON cf_scan_events (collection_id, barcode);
CREATE INDEX IF NOT EXISTS cf_scan_events_collection_received_idx
  ON cf_scan_events (collection_id, received_at);
CREATE INDEX IF NOT EXISTS cf_participants_collection_status_idx
  ON cf_participants (collection_id, status);

CREATE OR REPLACE VIEW cf_collection_totals AS
SELECT collection_id, barcode, SUM(quantity) AS total_quantity,
       COUNT(*) AS scan_count, MAX(received_at) AS last_received_at
FROM cf_scan_events
GROUP BY collection_id, barcode;

CREATE OR REPLACE FUNCTION cf_touch_collection_updated_at()
RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS cf_collections_touch_updated_at ON cf_collections;
CREATE TRIGGER cf_collections_touch_updated_at
BEFORE UPDATE ON cf_collections
FOR EACH ROW EXECUTE FUNCTION cf_touch_collection_updated_at();
