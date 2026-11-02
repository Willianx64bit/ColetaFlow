-- Store the owner's confirmed spreadsheet mapping for safe consolidated export.
ALTER TABLE cf_collections
  ADD COLUMN IF NOT EXISTS template_mapping jsonb;
