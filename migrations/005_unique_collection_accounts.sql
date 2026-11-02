-- A collection may contain users with the same display name.
-- A signed-in account has only one participant row per collection and can rejoin safely.
BEGIN;
ALTER TABLE cf_participants DROP CONSTRAINT IF EXISTS cf_participants_collection_id_display_name_key;
CREATE UNIQUE INDEX IF NOT EXISTS cf_participants_collection_user_key
  ON cf_participants(collection_id,user_id) WHERE user_id IS NOT NULL;
COMMIT;
